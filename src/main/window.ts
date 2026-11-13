import { join } from 'node:path'
import { BrowserWindow, type WebContents } from 'electron'
import { APP_NAME, WINDOW_DEFAULTS } from '@shared/constants.js'
import { IpcChannel, type PushIpcChannel } from '@shared/ipc-channels.js'
import { pushContracts, type MenuSnapshot, type PushPayload } from '@shared/ipc.js'
import type { MenuCommandPayload } from '@shared/api.js'
import { DiscardChoice, MenuCommand } from '@shared/types.js'
import { quitQuestion, type DirtyTab, type QuitQuestion } from '@services/tabs/quit.js'
import { confirmDiscardChanges, confirmDiscardSeveral } from './dialogs.js'
import { installContextMenu } from './context-menu.js'
import { SECURE_WEB_PREFERENCES } from './security-policy.js'
import { applyNavigationPolicy } from './security.js'
import { pathKey } from './fs/paths.js'

interface WindowState {
  dirty: readonly DirtyTab[]
  /** By `pathKey`: the files of its tabs. */
  paths: ReadonlySet<string>
  bypassGuard: boolean
  commands: MenuSnapshot
}

const HOME_SNAPSHOT: MenuSnapshot = { activeKind: null, editable: false, checked: [], disabled: [] }

const states = new WeakMap<BrowserWindow, WindowState>()

function stateOf(window: BrowserWindow): WindowState {
  const existing = states.get(window)
  if (existing !== undefined) return existing
  const created: WindowState = {
    dirty: [],
    paths: new Set(),
    bypassGuard: false,
    commands: HOME_SNAPSHOT,
  }
  states.set(window, created)
  return created
}

/** Absent in the production build. */
export function devServerUrl(): string | null {
  return process.env['ELECTRON_RENDERER_URL'] ?? null
}

/** Returns whether the menu has to be rebuilt. */
export function updateWindowState(
  window: BrowserWindow,
  shown: {
    title: string
    isDirty: boolean
    dirty: readonly DirtyTab[]
    commands: MenuSnapshot
    paths: readonly string[]
  },
): boolean {
  const state = stateOf(window)
  state.dirty = shown.dirty
  state.paths = new Set(shown.paths.map((path) => pathKey(path)))
  window.setTitle(`${shown.isDirty ? '• ' : ''}${shown.title} — ${APP_NAME}`)
  window.setDocumentEdited(shown.dirty.length > 0)

  const menuChanged = JSON.stringify(shown.commands) !== JSON.stringify(state.commands)
  state.commands = shown.commands
  return menuChanged
}

/** Another window with the file in a tab: reading it again would swap the package under that tab. */
export function windowHolding(path: string, asking: BrowserWindow): BrowserWindow | null {
  const key = pathKey(path)
  return (
    BrowserWindow.getAllWindows().find(
      (window) => window !== asking && !window.isDestroyed() && stateOf(window).paths.has(key),
    ) ?? null
  )
}

export function bringToFront(window: BrowserWindow): void {
  if (window.isMinimized()) window.restore()
  window.focus()
}

export function menuSnapshotOf(window: BrowserWindow | null): MenuSnapshot {
  return window === null ? HOME_SNAPSHOT : stateOf(window).commands
}

/** The renderer has already decided what to do. */
export function closeWithoutGuard(window: BrowserWindow): void {
  stateOf(window).bypassGuard = true
  window.close()
}

export function sendMenuCommand(window: BrowserWindow, payload: MenuCommandPayload): void {
  sendPush(window.webContents, IpcChannel.MenuCommand, payload)
}

/**
 * Validating the output catches a new field forgotten in the schema, which the renderer would drop
 * silently.
 */
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

function askBeforeClosing(
  window: BrowserWindow,
  question: Exclude<QuitQuestion, { kind: 'none' }>,
): Promise<DiscardChoice> {
  return question.kind === 'one'
    ? confirmDiscardChanges(window, question.label)
    : confirmDiscardSeveral(window, question.labels)
}

/** In main, so the prompt appears even if the renderer hangs. */
function installCloseGuard(window: BrowserWindow): void {
  window.on('close', (event) => {
    const state = stateOf(window)
    const question = quitQuestion(state.dirty)
    if (state.bypassGuard || question.kind === 'none') return

    event.preventDefault()

    void askBeforeClosing(window, question).then((choice) => {
      if (choice === DiscardChoice.Cancel) return

      if (choice === DiscardChoice.Discard) {
        closeWithoutGuard(window)
        return
      }

      // Only the renderer knows the content: it saves every tab and closes through
      // `closeWithoutGuard`; one cancel keeps the window.
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
    // No white flash: shown only after the renderer finished painting.
    show: false,
    backgroundColor: '#f6f7f9',
    webPreferences: {
      ...SECURE_WEB_PREFERENCES,
      // .cjs: sandboxed preloads do not support ESM.
      preload: join(import.meta.dirname, '../preload/index.cjs'),
    },
  })

  const url = devServerUrl()
  applyNavigationPolicy(window.webContents, url)
  installContextMenu(window.webContents)
  installCloseGuard(window)

  window.once('ready-to-show', () => window.show())

  if (url !== null) {
    // Vite serves pre-bundled dependencies as immutable for a year: a copy kept on disk from an
    // earlier run asks for chunks the new optimization no longer has, and the sheet grid stays blank.
    void window.webContents.session.clearCache().then(() => window.loadURL(url))
  } else {
    void window.loadFile(join(import.meta.dirname, '../renderer/index.html'))
  }

  return window
}
