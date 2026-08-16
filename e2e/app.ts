/**
 * Como subir o aplicativo num teste.
 *
 * Duas coisas precisam ser isoladas, senão os testes conversam entre si e com a
 * instalação real de quem está desenvolvendo:
 *
 * - **`userData` próprio por sessão de teste**: é onde moram os recentes e o
 *   rascunho de recuperação. Sem isolar, um teste apagaria os recentes do
 *   usuário — e a trava de instância única faria o segundo Electron desistir de
 *   subir porque o primeiro já tinha o mesmo diretório.
 * - **diálogos nativos**: `showOpenDialog` abre uma janela do sistema que nenhum
 *   teste consegue clicar. Eles são trocados por respostas fixas dentro do
 *   processo main, que é o único lugar onde isso é possível.
 */

import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { _electron as electron, type ElectronApplication, type Page } from '@playwright/test'

export interface Session {
  readonly app: ElectronApplication
  readonly window: Page
  /** Pasta de dados desta sessão. Sobrevive a um relaunch, de propósito. */
  readonly userData: string
  close: () => Promise<void>
  /** Mata o processo sem aviso, como uma queda de verdade. */
  crash: () => Promise<void>
}

/**
 * O executável empacotado, quando se quer testar o que de fato é distribuído.
 *
 * Sem a variável, os testes rodam sobre `out/` com o Electron de
 * desenvolvimento. Com ela — `LIBREVIA_E2E_BINARY=release/linux-unpacked/librevia`
 * — a mesma suíte roda contra o aplicativo instalado, que é onde caminhos de
 * recurso, asar e o binário do sidecar têm outra localização. É a diferença
 * entre "os testes passam" e "o instalador funciona".
 */
const packaged = process.env['LIBREVIA_E2E_BINARY']

export async function launch(options: { userData?: string } = {}): Promise<Session> {
  const userData = options.userData ?? (await mkdtemp(join(tmpdir(), 'librevia-e2e-')))

  const app = await electron.launch({
    ...(packaged === undefined || packaged === ''
      ? { args: [resolve('out/main/index.js'), `--user-data-dir=${userData}`] }
      : { executablePath: resolve(packaged), args: [`--user-data-dir=${userData}`] }),
    env: { ...process.env, NODE_ENV: 'production' },
  })

  const window = await app.firstWindow()
  await window.waitForLoadState('domcontentloaded')

  return {
    app,
    window,
    userData,
    // Encerramento à força de propósito: um teste que deixou trabalho não
    // salvo faria o aplicativo abrir o aviso nativo de descarte, e ninguém
    // clicaria nele. Nenhum teste aqui verifica saída limpa.
    close: async () => {
      app.process().kill('SIGKILL')
      await app.waitForEvent('close').catch(() => undefined)
      if (options.userData === undefined) await rm(userData, { recursive: true, force: true })
    },
    crash: async () => {
      // SIGKILL não roda nenhum handler de saída: é a diferença entre "fechou" e
      // "caiu", e a recuperação só vale para o segundo caso.
      app.process().kill('SIGKILL')
      await app.waitForEvent('close').catch(() => undefined)
    },
  }
}

/**
 * Troca os diálogos nativos por respostas fixas.
 *
 * Roda no processo main porque é lá que o `dialog` mora — o renderer não o
 * alcança, e é exatamente essa a garantia que o `contextIsolation` dá.
 */
export async function stubDialogs(
  app: ElectronApplication,
  answers: { open?: string; save?: string; messageBox?: number },
): Promise<void> {
  await app.evaluate(async ({ dialog }, fixed) => {
    const { open, save, messageBox } = fixed
    if (open !== undefined) {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [open] })
    }
    if (save !== undefined) {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: save })
    }
    if (messageBox !== undefined) {
      dialog.showMessageBox = async () => ({ response: messageBox, checkboxChecked: false })
    }
  }, answers)
}

/** Dispara um comando do menu nativo, que é como o aplicativo é operado. */
export async function menu(session: Session, command: string): Promise<void> {
  await session.app.evaluate(({ BrowserWindow }, name) => {
    BrowserWindow.getAllWindows()[0]?.webContents.send('menu:command', { command: name })
  }, command)
}
