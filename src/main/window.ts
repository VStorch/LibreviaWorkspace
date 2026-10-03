import { join } from 'node:path'
import { BrowserWindow, type WebContents } from 'electron'
import { APP_NAME, WINDOW_DEFAULTS } from '@shared/constants.js'
import { IpcChannel, type PushIpcChannel } from '@shared/ipc-channels.js'
import { pushContracts, type PushPayload } from '@shared/ipc.js'
import type { MenuCommandPayload } from '@shared/api.js'
import { DiscardChoice, MenuCommand } from '@shared/types.js'
import { confirmDiscardChanges } from './dialogs.js'
import { installContextMenu } from './context-menu.js'
import { t } from './i18n.js'
import { SECURE_WEB_PREFERENCES } from './security-policy.js'
import { applyNavigationPolicy } from './security.js'

interface WindowState {
  isDirty: boolean
  fileLabel: string
  bypassGuard: boolean
}

const states = new WeakMap<BrowserWindow, WindowState>()

function stateOf(window: BrowserWindow): WindowState {
  const existing = states.get(window)
  if (existing !== undefined) return existing
  const created: WindowState = { isDirty: false, fileLabel: t('shell.file.untitled'), bypassGuard: false }
  states.set(window, created)
  return created
}

/** URL do servidor do Vite em desenvolvimento; ausente na build de produção. */
export function devServerUrl(): string | null {
  return process.env['ELECTRON_RENDERER_URL'] ?? null
}

export function updateWindowState(window: BrowserWindow, title: string, isDirty: boolean): void {
  const state = stateOf(window)
  state.isDirty = isDirty
  state.fileLabel = title
  window.setTitle(`${isDirty ? '• ' : ''}${title} — ${APP_NAME}`)
  window.setDocumentEdited(isDirty)
}

/** O renderer já resolveu o que fazer. */
export function closeWithoutGuard(window: BrowserWindow): void {
  stateOf(window).bypassGuard = true
  window.close()
}

export function sendMenuCommand(window: BrowserWindow, payload: MenuCommandPayload): void {
  sendPush(window.webContents, IpcChannel.MenuCommand, payload)
}

/** Validar a saída pega o campo novo esquecido no schema, que o renderer descartaria em silêncio. */
export function sendPush<C extends PushIpcChannel>(
  contents: WebContents,
  channel: C,
  payload: PushPayload<C>,
): void {
  contents.send(channel, pushContracts[channel].parse(payload))
}

export function broadcastPush<C extends PushIpcChannel>(channel: C, payload: PushPayload<C>): void {
  for (const window of BrowserWindow.getAllWindows()) sendPush(window.webContents, channel, payload)
}

/** No main, para o aviso aparecer mesmo que o renderer trave. */
function installCloseGuard(window: BrowserWindow): void {
  window.on('close', (event) => {
    const state = stateOf(window)
    if (state.bypassGuard || !state.isDirty) return

    event.preventDefault()

    void confirmDiscardChanges(window, state.fileLabel).then((choice) => {
      if (choice === DiscardChoice.Cancel) return

      if (choice === DiscardChoice.Discard) {
        closeWithoutGuard(window)
        return
      }

      // Só o renderer sabe o conteúdo: ele grava e fecha por `closeWithoutGuard`.
      sendMenuCommand(window, { command: MenuCommand.SaveAndExit })
    })
  })
}

export function createMainWindow(): BrowserWindow {
  const window = new BrowserWindow({
    width: WINDOW_DEFAULTS.width,
    height: WINDOW_DEFAULTS.height,
    minWidth: WINDOW_DEFAULTS.minWidth,
    minHeight: WINDOW_DEFAULTS.minHeight,
    title: APP_NAME,
    // Sem o flash branco: só aparece quando o renderer terminou de pintar.
    show: false,
    backgroundColor: '#f6f7f9',
    webPreferences: {
      ...SECURE_WEB_PREFERENCES,
      // .cjs: preloads sandboxed não suportam ESM.
      preload: join(import.meta.dirname, '../preload/index.cjs'),
    },
  })

  const url = devServerUrl()
  applyNavigationPolicy(window.webContents, url)
  installContextMenu(window.webContents)
  installCloseGuard(window)

  window.once('ready-to-show', () => window.show())

  if (url !== null) {
    void window.loadURL(url)
  } else {
    void window.loadFile(join(import.meta.dirname, '../renderer/index.html'))
  }

  return window
}
