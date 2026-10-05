/**
 * How to launch the app in a test, isolated from the developer's installation:
 *
 * - its own `userData` per session, where recent files and the draft live, also because of the
 *   single-instance lock;
 * - native dialogs replaced by fixed answers in main, because no test clicks a system window.
 */

import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { _electron as electron, type ElectronApplication, type Page } from '@playwright/test'

export interface Session {
  readonly app: ElectronApplication
  readonly window: Page
  /** This session's data folder. Survives a relaunch, on purpose. */
  readonly userData: string
  close: () => Promise<void>
  /** Kills the process without warning, like a real crash. */
  crash: () => Promise<void>
}

/**
 * The packaged executable (`LIBREVIA_E2E_BINARY=release/linux-unpacked/librevia`), where resources,
 * asar and sidecar live elsewhere. Without the variable, `out/` with the development Electron.
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
    // Closes without the changes guard, but through Electron: killing only main leaves subprocesses
    // holding profile files open on Windows.
    close: async () => {
      const closed = app.waitForEvent('close')
      await app.evaluate(({ app }) => app.exit(0)).catch(() => undefined)
      await closed
      if (options.userData === undefined)
        await rm(userData, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 })
    },
    crash: async () => {
      // SIGKILL runs no exit handler: a crash, not a close.
      app.process().kill('SIGKILL')
      await app.waitForEvent('close').catch(() => undefined)
    },
  }
}

/** Replaces native dialogs with fixed answers, in main, where `dialog` lives. */
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
 * Fires a native menu command. Chromium serves input before IPC: one round through the page's task
 * queue lets the command run before the next key.
 */
export async function menu(session: Session, command: string): Promise<void> {
  await session.app.evaluate(({ BrowserWindow }, name) => {
    BrowserWindow.getAllWindows()[0]?.webContents.send('menu:command', { command: name })
  }, command)
  await session.window.evaluate(() => new Promise<void>((resolve) => setTimeout(resolve, 0)))
}
