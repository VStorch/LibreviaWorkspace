/**
 * How to launch the app in a test, isolated from the developer's installation:
 *
 * - its own `userData` per session, where recent files and the draft live, also because of the
 *   single-instance lock;
 * - native dialogs replaced by fixed answers in main, because no test clicks a system window.
 */

import { access, mkdtemp, rm, writeFile } from 'node:fs/promises'
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

/**
 * A fresh profile already answered the save format question: otherwise the first save of a new
 * file would stop at a native box no test clicks. `askSaveFormat` keeps it unanswered.
 */
async function seedProfile(userData: string, askSaveFormat: boolean): Promise<void> {
  const file = join(userData, 'preferences.json')
  if (
    askSaveFormat ||
    (await access(file).then(
      () => true,
      () => false,
    ))
  )
    return
  await writeFile(file, JSON.stringify({ preferences: { saveFormat: 'librevia' } }))
}

export async function launch(
  options: { userData?: string; file?: string; askSaveFormat?: boolean } = {},
): Promise<Session> {
  const userData = options.userData ?? (await mkdtemp(join(tmpdir(), 'librevia-e2e-')))
  await seedProfile(userData, options.askSaveFormat === true)
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
      // SIGKILL runs no exit handler: a crash, not a close. By the pid main reports: on Windows
      // `app.process()` is a wrapper, and killing it would leave the app holding its instance lock.
      const closed = app.waitForEvent('close').catch(() => undefined)
      process.kill(await app.evaluate(() => process.pid), 'SIGKILL')
      await closed
    },
  }
}

/** Replaces native dialogs with fixed answers, in main, where `dialog` lives. */
export async function stubDialogs(
  app: ElectronApplication,
  answers: { open?: string | string[]; save?: string; messageBox?: number },
): Promise<void> {
  await app.evaluate(async ({ dialog }, fixed) => {
    const { open, save, messageBox } = fixed
    if (open !== undefined) {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [open].flat() })
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
export async function menu(session: Session, command: string, windowIndex = 0): Promise<void> {
  await session.app.evaluate(
    ({ BrowserWindow }, [name, index]) => {
      const byAge = BrowserWindow.getAllWindows().sort((a, b) => a.id - b.id)
      byAge[index]?.webContents.send('menu:command', { command: name })
    },
    [command, windowIndex] as const,
  )
  const page = session.app.windows()[windowIndex] ?? session.window
  await page.evaluate(() => new Promise<void>((resolve) => setTimeout(resolve, 0)))
}

/** By the labels, as a click with the mouse: `menu()` skips the item and goes straight to the page. */
export async function clickMenuItem(session: Session, menuLabel: string, itemLabel: string): Promise<void> {
  await session.app.evaluate(
    ({ Menu }, [top, wanted]) => {
      const submenu = Menu.getApplicationMenu()?.items.find((entry) => entry.label === top)?.submenu
      const item = submenu?.items.find((entry) => entry.label === wanted)
      if (item === undefined) throw new Error(`no menu item ${top} > ${wanted}`)
      item.click()
    },
    [menuLabel, itemLabel] as const,
  )
}

/**
 * For a key pressed right after a click: until the next frames, the click may still move the selection
 * and the key would land elsewhere, faster than anyone types.
 */
export async function selectionSettled(session: Session): Promise<void> {
  await session.window.evaluate(
    () =>
      new Promise<void>((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(resolve, 0)))
      }),
  )
}
