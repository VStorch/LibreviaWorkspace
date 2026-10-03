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

// Sandbox para todo renderer, inclusive a janela oculta de impressão. Antes de `app.whenReady()`.
app.enableSandbox()
app.setName(APP_NAME)

// Também antes do `whenReady`: o esquema só ganha privilégio enquanto o Chromium monta a lista.
registerFontScheme()

// O dicionário antes de tudo, síncrono: dentro do `whenReady` o corretor já
// teria começado a baixar o dele, que sobrescreveria o embutido.
installBundledDictionary()

// Uma instância só: duas editando o mesmo arquivo perderiam dados.
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

    // O menu desenha as marcas pelas preferências, inclusive as mudadas pela barra.
    onPreferencesChanged(() => {
      void refreshMenu()
    })
    await refreshMenu()

    createMainWindow()

    // Depois da janela: o aplicativo não espera o sidecar para aparecer.
    void checkSidecarHealth()

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createMainWindow()
    })
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })

  // `will-quit`: no macOS o app segue vivo sem janela.
  app.on('will-quit', () => {
    disposeSidecar()
    // Ignorar é para a sessão; adicionar ao dicionário é para sempre.
    forgetSessionWords(session.defaultSession)
  })
}
