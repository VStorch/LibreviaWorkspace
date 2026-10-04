/**
 * Como subir o aplicativo num teste, isolado da instalação de quem desenvolve:
 *
 * - `userData` próprio por sessão, onde moram os recentes e o rascunho, também
 *   por causa da trava de instância única;
 * - diálogos nativos trocados por respostas fixas no processo main, porque nenhum
 *   teste clica numa janela do sistema.
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
 * O executável empacotado (`LIBREVIA_E2E_BINARY=release/linux-unpacked/librevia`),
 * onde recursos, asar e sidecar moram noutro lugar. Sem a variável, `out/` com o
 * Electron de desenvolvimento.
 */
const packaged = process.env['LIBREVIA_E2E_BINARY']

export async function launch(options: { userData?: string; file?: string } = {}): Promise<Session> {
  const userData = options.userData ?? (await mkdtemp(join(tmpdir(), 'librevia-e2e-')))
  const fileArgs = options.file === undefined ? [] : [options.file]

  const app = await electron.launch({
    ...(packaged === undefined || packaged === ''
      ? { args: [resolve('out/main/index.js'), `--user-data-dir=${userData}`, ...fileArgs] }
      : { executablePath: resolve(packaged), args: [`--user-data-dir=${userData}`, ...fileArgs] }),
    env: { ...process.env, NODE_ENV: 'production' },
  })

  const window = await app.firstWindow()
  await window.waitForLoadState('domcontentloaded')

  return {
    app,
    window,
    userData,
    // Encerra sem o guarda de alterações, mas pelo Electron: matar só o main
    // deixa subprocessos com arquivos do perfil abertos no Windows.
    close: async () => {
      const closed = app.waitForEvent('close')
      await app.evaluate(({ app }) => app.exit(0)).catch(() => undefined)
      await closed
      if (options.userData === undefined)
        await rm(userData, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 })
    },
    crash: async () => {
      // SIGKILL não roda handler de saída: é queda, e não fechamento.
      app.process().kill('SIGKILL')
      await app.waitForEvent('close').catch(() => undefined)
    },
  }
}

/** Troca os diálogos nativos por respostas fixas, no main, onde o `dialog` mora. */
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

/**
 * Dispara um comando do menu nativo. O Chromium atende a entrada antes do IPC: uma
 * volta pela fila de tarefas da página deixa o comando rodar antes da tecla seguinte.
 */
export async function menu(session: Session, command: string): Promise<void> {
  await session.app.evaluate(({ BrowserWindow }, name) => {
    BrowserWindow.getAllWindows()[0]?.webContents.send('menu:command', { command: name })
  }, command)
  await session.window.evaluate(() => new Promise<void>((resolve) => setTimeout(resolve, 0)))
}
