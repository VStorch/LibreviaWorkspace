import { mkdtemp, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithTrackedChange } from './fixtures.js'
import { gridReady } from './sheet-grid.js'

const TAB_BAR = 'Arquivos abertos'

async function draftCount(userData: string): Promise<number> {
  const files = await readdir(join(userData, 'recovery', 'drafts')).catch(() => [])
  return files.length
}

test.describe('one file per window', () => {
  let session: Session
  let userData: string
  let folder: string

  test.beforeEach(async () => {
    userData = await mkdtemp(join(tmpdir(), 'librevia-windows-'))
    folder = await mkdtemp(join(tmpdir(), 'librevia-windows-files-'))
    await writeFile(
      join(userData, 'preferences.json'),
      JSON.stringify({ preferences: { saveFormat: 'librevia', openMode: 'windows' } }),
    )
    session = await launch({ userData })
  })

  test.afterEach(async () => {
    await session.close()
    await rm(userData, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 })
    await rm(folder, { recursive: true, force: true })
  })

  async function nextWindow(command: string, from: number): Promise<Page> {
    const opened = session.app.waitForEvent('window')
    await menu(session, command, from)
    const page = await opened
    await page.waitForLoadState('domcontentloaded')
    return page
  }

  test('a file opened from a window that shows one gets its own, and a file already open shows its window', async () => {
    const first = session.window
    await menu(session, 'new-document')
    await expect(first.locator('.ProseMirror')).toBeFocused()
    await first.keyboard.type('Primeiro')
    await expect(first.getByRole('navigation', { name: TAB_BAR })).toHaveCount(0)

    const second = await nextWindow('new-spreadsheet', 0)
    await gridReady(second)
    await expect(first.locator('.ProseMirror')).toHaveText('Primeiro')
    await expect(second.getByRole('navigation', { name: TAB_BAR })).toHaveCount(0)

    const file = join(folder, 'ata.docx')
    await writeFile(file, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: file })
    const third = await nextWindow('open', 0)
    await expect(third.locator('.ProseMirror')).toContainText('Ata da reunião de terça.')

    await menu(session, 'open', 1)
    await expect(second.locator('revo-grid')).toBeVisible()
    await expect.poll(() => session.app.windows().length).toBe(3)
    await expect(second.locator('.ProseMirror')).toHaveCount(0)
  })

  test('a new window does not offer the drafts of the tabs another window still has open', async () => {
    const first = session.window
    await menu(session, 'new-document')
    await expect(first.locator('.ProseMirror')).toBeFocused()
    await first.keyboard.type('Ainda aberto')
    await expect.poll(() => draftCount(userData), { timeout: 15_000 }).toBe(1)

    const second = await nextWindow('new-spreadsheet', 0)
    await gridReady(second)
    await second.waitForTimeout(500)
    await expect(second.getByRole('button', { name: 'Recuperar' })).toHaveCount(0)
  })
})
