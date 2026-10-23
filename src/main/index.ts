import { app, BrowserWindow, session } from 'electron'
import { aiSettings } from './ai/settings.js'
import { APP_NAME } from '@shared/constants.js'
import { registerAiHandlers } from './ipc/ai.js'
import { registerMcpHandlers } from './ipc/mcp.js'
import { registerEditingHandlers } from './ipc/editing.js'
import { registerFileHandlers } from './ipc/file.js'
import { registerExportHandlers } from './ipc/export.js'
import { registerPdfHandlers } from './ipc/pdf.js'
import { registerPrivacyHandlers } from './ipc/privacy.js'
import { registerPrintHandlers } from './ipc/print.js'
import { registerRecoveryHandlers } from './ipc/recovery.js'
import { registerTemplateHandlers } from './ipc/templates.js'
import { registerWindowHandlers } from './ipc/window.js'
import { useRecoveryFolder } from './fs/recovery.js'
import { refreshMenu } from './menu.js'
import { serveFonts } from './fonts.js'
import { serveViewer } from './pdf-scheme.js'
import { registerAppSchemes } from './schemes.js'
import { applyStoredPreferences, onPreferencesChanged } from './preferences.js'
import { stopAllAgents } from './ai/acp.js'
import { attachMcpWindow, disposeMcp, mcpHost } from './ai/mcp.js'
import { runMcpBridge } from './ai/mcp-bridge.js'
import { isMcpBridge } from './ai/mcp-paths.js'
import { applySessionPolicy } from './security.js'
import { forgetSessionWords, installBundledDictionary, serveDictionary } from './spellcheck.js'
import { checkSidecarHealth, disposeSidecar } from './sidecar/index.js'
import { createMainWindow as createWindow, devServerUrl } from './window.js'
import { fileFromArguments, requestExternalFile } from './external-files.js'
import { watchProcessFailures, watchWindowFailures } from './failures.js'
import { logInfo, logWarning, useLogFolder } from './log.js'

// The bridge of a client of the local MCP server: before anything of the app, which it never starts.
if (isMcpBridge(process.argv)) runMcpBridge()
else startApp()

function startApp(): void {
  // Sandbox for every renderer, including the hidden print window. Before `app.whenReady()`.
  app.enableSandbox()
  app.setName(APP_NAME)

  useLogFolder(app.getPath('logs'))
  watchProcessFailures()

  // Also before `whenReady`: a scheme only gets privileges while Chromium builds the list.
  registerAppSchemes()

  // The dictionary first, synchronously: inside `whenReady` the spellchecker would already have
  // started downloading its own, which would overwrite the bundled one.
  installBundledDictionary()

  // A single instance: two editing the same file would lose data.
  if (!app.requestSingleInstanceLock()) {
    app.quit()
    return
  }
  const initialFile = fileFromArguments(process.argv.slice(app.isPackaged ? 1 : 2), process.cwd())
  if (initialFile !== undefined) requestExternalFile(initialFile)

  app.on('second-instance', (_event, argv, workingDirectory) => {
    const path = fileFromArguments(argv.slice(app.isPackaged ? 1 : 2), workingDirectory)
    if (path !== undefined) requestExternalFile(path)
    const [existing] = BrowserWindow.getAllWindows()
    if (existing === undefined) return
    if (existing.isMinimized()) existing.restore()
    existing.focus()
  })

  void app.whenReady().then(onReady)

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })

  // `will-quit`: on macOS the app stays alive without windows.
  app.on('will-quit', () => {
    disposeSidecar()
    stopAllAgents()
    disposeMcp()
    // Ignoring lasts for the session; adding to the dictionary is forever.
    forgetSessionWords(session.defaultSession)
  })
}

function registerHandlers(): void {
  registerEditingHandlers()
  registerFileHandlers()
  registerPrintHandlers()
  registerPdfHandlers()
  registerPrivacyHandlers()
  registerExportHandlers()
  registerRecoveryHandlers()
  registerTemplateHandlers()
  registerWindowHandlers()
  registerAiHandlers()
  registerMcpHandlers()
}

/** The local MCP server answers through the window that shows the document. */
function createMainWindow(): void {
  const window = createWindow()
  watchWindowFailures(window)
  attachMcpWindow(window.webContents)
}

async function onReady(): Promise<void> {
  applySessionPolicy(session.defaultSession, devServerUrl() === null ? 'production' : 'development')
  serveFonts()
  serveViewer()
  serveDictionary()
  applyStoredPreferences()

  useRecoveryFolder(app.getPath('userData'))
  logInfo(
    'app',
    `${APP_NAME} ${app.getVersion()}, ${process.platform} ${process.arch}, Electron ${process.versions.electron}`,
  )

  registerHandlers()

  // The menu draws its check marks from the preferences, including those changed from the
  // toolbar.
  onPreferencesChanged(() => {
    void refreshMenu()
  })
  await refreshMenu()

  createMainWindow()

  // After the window: the app does not wait for the sidecar to appear.
  void checkSidecarHealth()
  // A new key on each opening, when the person allowed external programs.
  if (aiSettings().preferences().mcpServer) {
    mcpHost()
      .setExternal(true)
      .catch((cause: unknown) => logWarning('mcp', 'the local server did not start:', String(cause)))
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow()
  })
}
