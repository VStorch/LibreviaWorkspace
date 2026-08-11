import { join } from 'node:path'
import { BrowserWindow } from 'electron'
import { APP_NAME, WINDOW_DEFAULTS } from '@shared/constants.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { MenuCommandPayload } from '@shared/api.js'
import { DiscardChoice, MenuCommand } from '@shared/types.js'
import { confirmDiscardChanges } from './dialogs.js'
import { SECURE_WEB_PREFERENCES } from './security-policy.js'
import { applyNavigationPolicy } from './security.js'

interface WindowState {
  isDirty: boolean
  fileLabel: string
  /** Ligado quando o fechamento já foi decidido e não deve ser interceptado de novo. */
  bypassGuard: boolean
}

const states = new WeakMap<BrowserWindow, WindowState>()

function stateOf(window: BrowserWindow): WindowState {
  const existing = states.get(window)
  if (existing !== undefined) return existing
  const created: WindowState = { isDirty: false, fileLabel: 'Sem título', bypassGuard: false }
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
  // No macOS a bolinha no botão de fechar é a convenção nativa.
  window.setDocumentEdited(isDirty)
}

/** Fecha sem passar pelo guarda — o renderer já resolveu o que fazer. */
export function closeWithoutGuard(window: BrowserWindow): void {
  stateOf(window).bypassGuard = true
  window.close()
}

export function sendMenuCommand(window: BrowserWindow, payload: MenuCommandPayload): void {
  window.webContents.send(IpcChannel.MenuCommand, payload)
}

/**
 * Guarda de fechamento.
 *
 * É a última linha contra perda de trabalho, e por isso mora no processo main:
 * mesmo que o renderer trave ou seja fechado pelo gerenciador de janelas, o
 * aviso aparece.
 */
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

      // "Salvar": só o renderer sabe o conteúdo atual. Ele grava e então
      // chama window.close() pela API, que passa por closeWithoutGuard.
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
    // Evita o flash branco: só mostramos quando o renderer terminou de pintar.
    show: false,
    backgroundColor: '#f6f7f9',
    webPreferences: {
      ...SECURE_WEB_PREFERENCES,
      // .cjs e não .mjs: preloads sandboxed não suportam ESM.
      preload: join(import.meta.dirname, '../preload/index.cjs'),
    },
  })

  const url = devServerUrl()
  applyNavigationPolicy(window.webContents, url)
  installCloseGuard(window)

  window.once('ready-to-show', () => window.show())

  if (url !== null) {
    void window.loadURL(url)
  } else {
    void window.loadFile(join(import.meta.dirname, '../renderer/index.html'))
  }

  return window
}
