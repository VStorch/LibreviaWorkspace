import { app, BrowserWindow, session } from 'electron'
import { APP_NAME } from '@shared/constants.js'
import { registerEditingHandlers } from './ipc/editing.js'
import { registerFileHandlers } from './ipc/file.js'
import { registerPrintHandlers } from './ipc/print.js'
import { registerRecoveryHandlers } from './ipc/recovery.js'
import { registerWindowHandlers } from './ipc/window.js'
import { useRecoveryFolder } from './fs/recovery.js'
import { refreshMenu } from './menu.js'
import { registerFontScheme, serveFonts } from './fonts.js'
import { applyStoredPreferences, onPreferencesChanged } from './preferences.js'
import { applySessionPolicy } from './security.js'
import { forgetSessionWords, installBundledDictionary, serveDictionary } from './spellcheck.js'
import { checkSidecarHealth, disposeSidecar } from './sidecar/index.js'
import { createMainWindow, devServerUrl } from './window.js'

// Sandbox para todo renderer, inclusive os que vierem depois (janela oculta de
// impressão, por exemplo). Precisa vir antes de `app.whenReady()`.
app.enableSandbox()
app.setName(APP_NAME)

// Também antes do `whenReady`: um esquema só ganha privilégio se for declarado
// enquanto o Chromium ainda está montando a lista.
registerFontScheme()

// E o dicionário de português antes de tudo, síncrono.
//
// Isto foi medido: feito dentro do `whenReady`, o corretor do Chromium já havia
// começado a baixar o dicionário — ele se inicializa quando a sessão padrão passa
// a existir, e a sessão passa a existir na primeira linha que a menciona. O
// arquivo baixado sobrescrevia o embutido, e o aplicativo ficava dependendo de
// rede sem ninguém notar. Ver src/main/spellcheck.ts.
installBundledDictionary()

// Uma instância só: duas instâncias editando o mesmo arquivo é caminho certo
// para perda de dados.
if (!app.requestSingleInstanceLock()) {
  app.quit()
} else {
  app.on('second-instance', () => {
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

    // A pasta chega por parâmetro para que o módulo de recuperação não dependa
    // do Electron — é o que permite testá-lo sem subir um aplicativo inteiro.
    useRecoveryFolder(app.getPath('userData'))

    registerEditingHandlers()
    registerFileHandlers()
    registerPrintHandlers()
    registerRecoveryHandlers()
    registerWindowHandlers()

    // O menu desenha as marcas de seleção a partir das preferências, então
    // precisa ser refeito quando alguma muda — inclusive quando quem a mudou foi
    // a barra de ferramentas.
    onPreferencesChanged(() => {
      void refreshMenu()
    })
    await refreshMenu()

    createMainWindow()

    // Depois da janela: o aplicativo não espera pelo serviço de formatos para
    // aparecer na tela.
    void checkSidecarHealth()

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createMainWindow()
    })
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })

  // `will-quit` e não `window-all-closed`: no macOS o app segue vivo sem janela,
  // e matar o sidecar ali deixaria o próximo documento sem serviço.
  app.on('will-quit', () => {
    disposeSidecar()
    // As palavras "ignoradas" desta sessão saem do dicionário do usuário: ignorar
    // é para agora, adicionar ao dicionário é para sempre.
    forgetSessionWords(session.defaultSession)
  })
}
