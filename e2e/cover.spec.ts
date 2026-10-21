import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithAnchoredTextBox, entryOf } from './fixtures.js'

/**
 * The cover is editable: in the manual template, title and subtitle live in positioned boxes,
 * outside the `contenteditable`. Text typed in the box becomes an attribute of the anchor block and
 * goes out in `w:txbxContent`, without losing the box.
 */
test.describe('editable cover', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-cover-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('the box title is typeable and returns to the file', async () => {
    const sourceFile = join(folder, 'capa.docx')
    const destinationFile = join(folder, 'salva.docx')
    await writeFile(sourceFile, await docxWithAnchoredTextBox())

    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')

    // The document does not open locked: the shape is no reason for a lock.
    await expect(session.window.locator('.readonly-banner')).toHaveCount(0)

    const box = session.window.locator('.paper-float--text').filter({ hasText: 'Título da capa' })
    await expect(box).toHaveCount(1)

    await box.click()
    await session.window.keyboard.press('ControlOrMeta+a')
    await session.window.keyboard.type('Título trocado')

    // Text leaves the box when it loses focus: only then does the block attribute change, so the
    // sheet is not redrawn under the cursor.
    await session.window.locator('.ProseMirror').click()
    await expect(session.window.locator('.paper-float--text')).toContainText('Título trocado')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const body = await entryOf(destinationFile, 'word/document.xml')
    expect(body).toContain('txbxContent')
    expect(body).toContain('Título trocado')
    expect(body).not.toContain('Título da capa')
  })
})
