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
test.describe('gravação cirúrgica', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-preservacao-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('abrir e salvar sem editar preserva a caixa de texto', async () => {
    const origem = join(folder, 'entrada.docx')
    const destino = join(folder, 'saida.docx')
    await writeFile(origem, await docxWithTextBox())

    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Título na caixa')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // The box is the sentinel: the editor shows the text inside, but cannot redraw the shape. If it
    // came back, the block's original XML was preserved, which is all this test needs to know.
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo).toContain('txbxContent')
    expect(corpo).toContain('Título na caixa')
  })

  test('documento que termina num título volta sem parágrafo vazio acrescentado', async () => {
    // Tiptap added an empty paragraph after the last heading, and saving without editing put a new
    // `<w:p/>` at the end of the file; the four corpus evidence documents end that way.
    const origem = join(folder, 'titulo-no-fim.docx')
    const destino = join(folder, 'titulo-no-fim-saida.docx')
    const final = '<w:p><w:pPr><w:pStyle w:val="Ttulo1"/></w:pPr><w:r><w:t>Título final</w:t></w:r></w:p>'
    await writeFile(origem, await docxWithNamedStyles(final))

    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Título final')
    await expect(session.window.locator('.ProseMirror > *').last()).toHaveText('Título final')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const paragrafos = (xml: string): number => (xml.match(/<w:p[ >/]/g) ?? []).length
    expect(paragrafos(await entryOf(destino, 'word/document.xml'))).toBe(
      paragrafos(await entryOf(origem, 'word/document.xml')),
    )
  })
})
