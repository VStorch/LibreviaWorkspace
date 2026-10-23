import { app, dialog, type BrowserWindow } from 'electron'
import { t } from './i18n.js'
import { logError, logFilePath, logWarning } from './log.js'
import { closeWithoutGuard } from './window.js'

const RELOAD_BUTTON = 0
const CLOSE_BUTTON = 1

let mainErrorShown = false

export function watchProcessFailures(): void {
  process.on('uncaughtException', (error) => {
    logError('main', 'uncaught exception:', error)
    showMainErrorOnce()
  })
  process.on('unhandledRejection', (reason) => logError('main', 'unhandled rejection:', reason))

  app.on('child-process-gone', (_event, details) => {
    if (details.reason === 'clean-exit') return
    logError(
      'process',
      `${details.type} gone (${details.reason}, code ${details.exitCode})`,
      details.name ?? '',
    )
  })
}

function showMainErrorOnce(): void {
  if (mainErrorShown || !app.isReady()) return
  mainErrorShown = true
  dialog.showErrorBox(t('shell.failure.main.title'), t('shell.failure.main.detail', { log: logLocation() }))
}

function logLocation(): string {
  return logFilePath() ?? '—'
}

export function watchWindowFailures(window: BrowserWindow): void {
  const contents = window.webContents

  contents.on('console-message', (event) => {
    if (event.level !== 'error') return
    logError('renderer', `${event.message} (${event.sourceId}:${event.lineNumber})`)
  })
  contents.on('preload-error', (_event, path, error) => logError('preload', path, error))
  contents.on('unresponsive', () => logWarning('window', 'the window stopped responding'))

  contents.on('render-process-gone', (_event, details) => {
    if (details.reason === 'clean-exit') return
    logError('window', `renderer gone (${details.reason}, code ${details.exitCode})`)
    if (!window.isDestroyed()) void offerReload(window)
  })
}

async function offerReload(window: BrowserWindow): Promise<void> {
  const { response } = await dialog.showMessageBox(window, {
    type: 'error',
    message: t('shell.failure.window.title'),
    detail: t('shell.failure.window.detail', { log: logLocation() }),
    buttons: [t('shell.failure.window.reload'), t('shell.failure.window.close')],
    defaultId: RELOAD_BUTTON,
    cancelId: CLOSE_BUTTON,
    noLink: true,
  })
  if (window.isDestroyed()) return
  if (response === RELOAD_BUTTON) void window.webContents.loadURL(window.webContents.getURL())
  else closeWithoutGuard(window)
}
