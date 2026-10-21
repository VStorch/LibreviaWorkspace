import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithHeaderTextBox, entryOf } from './fixtures.js'

/**
 * A header that is a shape group, like most of the corpus: the title in an anchored box. The box
 * comes back whole, because typing in it opens and closes paragraphs; the `PAGE` field does not, or
 * it would become a fixed number.
 */
test.describe('editable header box', () => {
  let session: Session
  let directory: string
  let destinationFile: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-band-'))
    destinationFile = join(directory, 'salva.docx')
    session = await launch()

    const sourceFile = join(directory, 'grupo.docx')
    await writeFile(sourceFile, await docxWithHeaderTextBox())
    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  test('the box title is typeable and returns to the file header', async () => {
    const box = session.window
      .locator('.paper-float--text')
      .filter({ hasText: 'EVIDÊNCIAS DO ROTEIRO' })
      .first()

    await box.click()
    await session.window.keyboard.press('ControlOrMeta+a')
    await session.window.keyboard.type('EVIDÊNCIAS DE HOMOLOGAÇÃO')
    await session.window.locator('.ProseMirror').click()

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const header = await entryOf(destinationFile, 'word/header1.xml')
    expect(header).toContain('EVIDÊNCIAS DE HOMOLOGAÇÃO')
    expect(header).not.toContain('EVIDÊNCIAS DO ROTEIRO')

    // The group stays whole, with the field this writer cannot generate.
    expect(header).toContain('PAGE')
    expect(header).toContain('fldChar')
  })

  test('the page number box does not receive the cursor', async () => {
    // What shows in it is this sheet's number; what is in the file is a field. Returning it as text
    // would make the header say "1" on every sheet, and it would only be noticed on the second.
    const number = session.window.locator('.paper-float--text').filter({ hasText: /^1$/ })
    await expect(number).toHaveCount(1)
    await expect(number).not.toHaveClass(/paper-float--edit/)
  })
})
