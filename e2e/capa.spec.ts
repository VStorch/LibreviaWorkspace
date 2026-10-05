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
test.describe('capa editável', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-capa-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('o título da caixa é digitável e volta para o arquivo', async () => {
    const origem = join(folder, 'capa.docx')
    const destino = join(folder, 'salva.docx')
    await writeFile(origem, await docxWithAnchoredTextBox())

    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')

    // The document does not open locked: the shape is no reason for a lock.
    await expect(session.window.locator('.readonly-banner')).toHaveCount(0)

    const caixa = session.window.locator('.paper-float--text').filter({ hasText: 'Título da capa' })
    await expect(caixa).toHaveCount(1)

    await caixa.click()
    await session.window.keyboard.press('ControlOrMeta+a')
    await session.window.keyboard.type('Título trocado')

    // Text leaves the box when it loses focus: only then does the block attribute change, so the
    // sheet is not redrawn under the cursor.
    await session.window.locator('.ProseMirror').click()
    await expect(session.window.locator('.paper-float--text')).toContainText('Título trocado')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo).toContain('txbxContent')
    expect(corpo).toContain('Título trocado')
    expect(corpo).not.toContain('Título da capa')
  })
})
