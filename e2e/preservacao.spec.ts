import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithNamedStyles, docxWithTextBox } from './fixtures.js'

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
    const corpo = await corpoDoDocumento(destino)
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
    expect(paragrafos(await corpoDoDocumento(destino))).toBe(paragrafos(await corpoDoDocumento(origem)))
  })
})

/** The content of `word/document.xml` inside the `.docx`, without unpacking to disk. */
async function corpoDoDocumento(caminho: string): Promise<string> {
  const { promisify } = await import('node:util')
  const { inflateRaw } = await import('node:zlib')
  const inflate = promisify(inflateRaw)
  const zip = await readFile(caminho)

  // A sweep of the zip local headers: enough to find a part by name, without a new test dependency.
  for (let i = 0; i + 30 <= zip.length; i++) {
    if (zip.readUInt32LE(i) !== 0x04034b50) continue

    const metodo = zip.readUInt16LE(i + 8)
    const comprimido = zip.readUInt32LE(i + 18)
    const original = zip.readUInt32LE(i + 22)
    const tamanhoNome = zip.readUInt16LE(i + 26)
    const extra = zip.readUInt16LE(i + 28)
    const nome = zip.subarray(i + 30, i + 30 + tamanhoNome).toString('utf8')
    if (nome !== 'word/document.xml') continue

    const fim = i + 30 + tamanhoNome + extra + (comprimido > 0 ? comprimido : original)
    const dados = zip.subarray(i + 30 + tamanhoNome + extra, fim)
    return (metodo === 0 ? dados : await inflate(dados)).toString('utf8')
  }

  throw new Error(`word/document.xml não encontrado em ${caminho}`)
}
