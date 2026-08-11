import { app, BrowserWindow, session } from 'electron'
import { APP_NAME } from '@shared/constants.js'
import { registerFileHandlers } from './ipc/file.js'
import { registerWindowHandlers } from './ipc/window.js'
import { refreshMenu } from './menu.js'
import { applySessionPolicy } from './security.js'
import { createMainWindow, devServerUrl } from './window.js'

// Sandbox para todo renderer, inclusive os que vierem depois (janela oculta de
// impressão, por exemplo). Precisa vir antes de `app.whenReady()`.
app.enableSandbox()
app.setName(APP_NAME)

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

    registerFileHandlers()
    registerWindowHandlers()
    await refreshMenu()

    createMainWindow()

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createMainWindow()
    })
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })
}
