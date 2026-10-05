import { app, BrowserWindow, session } from 'electron'
import { APP_NAME } from '@shared/constants.js'
import { registerEditingHandlers } from './ipc/editing.js'
import { registerFileHandlers } from './ipc/file.js'
import { registerExportHandlers } from './ipc/export.js'
import { registerPrintHandlers } from './ipc/print.js'
import { registerRecoveryHandlers } from './ipc/recovery.js'
import { registerTemplateHandlers } from './ipc/templates.js'
import { registerWindowHandlers } from './ipc/window.js'
import { useRecoveryFolder } from './fs/recovery.js'
import { refreshMenu } from './menu.js'
import { registerFontScheme, serveFonts } from './fonts.js'
import { applyStoredPreferences, onPreferencesChanged } from './preferences.js'
import { applySessionPolicy } from './security.js'
import { forgetSessionWords, installBundledDictionary, serveDictionary } from './spellcheck.js'
import { checkSidecarHealth, disposeSidecar } from './sidecar/index.js'
import { createMainWindow, devServerUrl } from './window.js'
import { docxFromArguments, requestExternalFile } from './external-files.js'

// Sandbox for every renderer, including the hidden print window. Before `app.whenReady()`.
app.enableSandbox()
app.setName(APP_NAME)

// Also before `whenReady`: the scheme only gets privileges while Chromium builds the list.
registerFontScheme()

// The dictionary first, synchronously: inside `whenReady` the spellchecker would already have
// started downloading its own, which would overwrite the bundled one.
installBundledDictionary()

// A single instance: two editing the same file would lose data.
if (!app.requestSingleInstanceLock()) {
  app.quit()
} else {
  const initialFile = docxFromArguments(process.argv.slice(app.isPackaged ? 1 : 2), process.cwd())
  if (initialFile !== undefined) requestExternalFile(initialFile)

  app.on('second-instance', (_event, argv, workingDirectory) => {
    const path = docxFromArguments(argv.slice(app.isPackaged ? 1 : 2), workingDirectory)
    if (path !== undefined) requestExternalFile(path)
    const [existing] = BrowserWindow.getAllWindows()
    if (existing === undefined) return
    if (existing.isMinimized()) existing.restore()
    existing.focus()
  })

  void app.whenReady().then(async () => {
    applySessionPolicy(session.defaultSession, devServerUrl() === null ? 'production' : 'development')
    serveFonts()
    serveDictionary()
    applyStoredPreferences()

    useRecoveryFolder(app.getPath('userData'))

    registerEditingHandlers()
    registerFileHandlers()
    registerPrintHandlers()
    registerExportHandlers()
    registerRecoveryHandlers()
    registerTemplateHandlers()
    registerWindowHandlers()

    // The menu draws its check marks from the preferences, including those changed from the
    // toolbar.
    onPreferencesChanged(() => {
      void refreshMenu()
    })
    await refreshMenu()

    createMainWindow()

    // After the window: the app does not wait for the sidecar to appear.
    void checkSidecarHealth()

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createMainWindow()
    })
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })

  // `will-quit`: on macOS the app stays alive without windows.
  app.on('will-quit', () => {
    disposeSidecar()
    // Ignoring lasts for the session; adding to the dictionary is forever.
    forgetSessionWords(session.defaultSession)
  })
}
