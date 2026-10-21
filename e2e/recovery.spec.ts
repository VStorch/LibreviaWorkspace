import { stat } from 'node:fs/promises'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/** Recovery proven against a real crash: `SIGKILL`, without exit handlers. */
test.describe('recovery after a crash', () => {
  let session: Session

  test.afterEach(async () => {
    await session.close()
  })

  test('restores what was on screen and writes to no file', async () => {
    session = await launch()

    await menu(session, 'new-document')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()

    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Ata da reunião de terça')

    // Autosave runs on a timer, every eight seconds.
    await expect
      .poll(async () => draftOnDisk(session), { timeout: 20_000, message: 'the draft was not written' })
      .toBe(true)

    const { userData } = session
    await session.crash()

    // Same data folder: that is what makes the second session find the first one's draft, as would
    // happen with the same user on the same machine.
    session = await launch({ userData })

    const banner = session.window.locator('.banner--recovery')
    await expect(banner).toBeVisible()
    await expect(banner).toContainText('trabalho não salvo')

    await banner.getByRole('button', { name: 'Recuperar' }).click()

    await expect(session.window.locator('.ProseMirror')).toContainText('Ata da reunião de terça')
    // Recovered content differs from what is on disk: marking it saved would let the user close the
    // window thinking everything was kept.
    await expect(session.window.locator('.statusbar__state')).toContainText('Não salvo')
  })

  test('discarding deletes the draft for good', async () => {
    session = await launch()

    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Rascunho a descartar')
    await expect.poll(async () => draftOnDisk(session), { timeout: 20_000 }).toBe(true)

    const { userData } = session
    await session.crash()

    session = await launch({ userData })
    await session.window.locator('.banner--recovery').getByRole('button', { name: 'Descartar' }).click()
    await expect(session.window.locator('.banner--recovery')).toBeHidden()

    expect(await draftOnDisk(session)).toBe(false)
  })

  test('a clean session shows no notice', async () => {
    // A warning that appears on every launch is one the user learns to close without reading.
    session = await launch()

    await session.window.waitForTimeout(2000)
    await expect(session.window.locator('.banner--recovery')).toBeHidden()
  })
})

/** Is the draft on disk? Checked by path, not through the API under test. */
async function draftOnDisk(session: Session): Promise<boolean> {
  try {
    await stat(join(session.userData, 'recuperacao', 'rascunho.json'))
    return true
  } catch {
    return false
  }
}
