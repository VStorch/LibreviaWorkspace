import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithNamedStyles, docxWithTextBox, entryOf } from './fixtures.js'

/**
 * Opening and saving without editing does not touch the file: the writer returns the original XML
 * of every block that did not change, through its `oid`. The path goes through ProseMirror, where
 * `oid` must be in the schema and shape differences do not count; the sidecar tests do not go
 * through the editor.
 */
test.describe('surgical save', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-preservation-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('opening and saving without editing preserves the text box', async () => {
    const sourceFile = join(folder, 'entrada.docx')
    const destinationFile = join(folder, 'saida.docx')
    await writeFile(sourceFile, await docxWithTextBox())

    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Título na caixa')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // The box is the sentinel: the editor shows the text inside, but cannot redraw the shape. If it
    // came back, the block's original XML was preserved, which is all this test needs to know.
    const body = await entryOf(destinationFile, 'word/document.xml')
    expect(body).toContain('txbxContent')
    expect(body).toContain('Título na caixa')
  })

  test('a document ending in a heading comes back without an added empty paragraph', async () => {
    // Tiptap added an empty paragraph after the last heading, and saving without editing put a new
    // `<w:p/>` at the end of the file; the four corpus evidence documents end that way.
    const sourceFile = join(folder, 'titulo-no-fim.docx')
    const destinationFile = join(folder, 'titulo-no-fim-saida.docx')
    const final = '<w:p><w:pPr><w:pStyle w:val="Ttulo1"/></w:pPr><w:r><w:t>Título final</w:t></w:r></w:p>'
    await writeFile(sourceFile, await docxWithNamedStyles(final))

    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Título final')
    await expect(session.window.locator('.ProseMirror > *').last()).toHaveText('Título final')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const paragraphs = (xml: string): number => (xml.match(/<w:p[ >/]/g) ?? []).length
    expect(paragraphs(await entryOf(destinationFile, 'word/document.xml'))).toBe(
      paragraphs(await entryOf(sourceFile, 'word/document.xml')),
    )
  })
})
