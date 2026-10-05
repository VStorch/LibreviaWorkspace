import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithFootnote, entryOf } from './fixtures.js'

/**
 * File → Export as → HTML…, Markdown… and ODT…. Exporting writes a new file: the document stays at
 * its path and without pending changes.
 */
test.describe('exportar como HTML, Markdown e ODT', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-exportar-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  const lido = (path: string) => () => readFile(path, 'utf8').catch(() => '')

  /** The ZIP entry names, from the central directory, in file order. */
  const entradas = (zip: Buffer): string[] => {
    const fim = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
    const nomes: string[] = []
    let i = zip.readUInt32LE(fim + 16)
    for (let n = zip.readUInt16LE(fim + 10); n > 0; n--) {
      const nome = zip.readUInt16LE(i + 28)
      nomes.push(zip.subarray(i + 46, i + 46 + nome).toString('utf8'))
      i += 46 + nome + zip.readUInt16LE(i + 30) + zip.readUInt16LE(i + 32)
    }
    return nomes
  }

  test('exporta o documento aberto sem trocar o caminho dele nem sujá-lo', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithFootnote())
    const antes = (await stat(origem)).mtimeMs
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Ata da reunião')

    const html = join(pasta, 'ata.html')
    await stubDialogs(session.app, { save: html })
    await menu(session, 'export-html')
    await expect.poll(lido(html), { timeout: 15_000 }).toContain('</html>')
    const pagina = await readFile(html, 'utf8')
    expect(pagina).toContain('<html lang="pt-BR">')
    expect(pagina).toContain('<title>ata</title>')
    expect(pagina).toContain('<p>Ata da reunião de terça.</p>')
    expect(pagina).toContain('<a href="#nota-rodape-1" id="ref-nota-rodape-1">1</a>')
    expect(pagina).toContain('Fonte: ata anterior.')
    expect(pagina).not.toMatch(/<script/i)

    const markdown = join(pasta, 'ata.md')
    await stubDialogs(session.app, { save: markdown })
    await menu(session, 'export-markdown')
    await expect.poll(lido(markdown), { timeout: 15_000 }).toContain('[^1]:')
    const texto = await readFile(markdown, 'utf8')
    expect(texto).toContain('Ata da reunião de terça.\n\nSegundo parágrafo, com uma nota.[^1]')
    expect(texto).toContain('[^1]: Fonte: ata anterior.')

    // The document is still the .docx, clean, and the source file intact.
    await expect(session.window).toHaveTitle(/^ata\.docx/)
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect((await stat(origem)).mtimeMs).toBe(antes)
  })

  test('exporta para ODT um pacote OpenDocument com as partes dele, sem sujar o documento', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithFootnote())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Ata da reunião')

    const odt = join(pasta, 'ata.odt')
    await stubDialogs(session.app, { save: odt })
    await menu(session, 'export-odt')
    await expect
      .poll(
        () =>
          stat(odt)
            .then((info) => info.size)
            .catch(() => 0),
        { timeout: 15_000 },
      )
      .toBeGreaterThan(0)

    // `mimetype` opens the package, stored, and all the parts are there.
    const pacote = await readFile(odt)
    expect(pacote.subarray(30, 38).toString()).toBe('mimetype')
    expect(pacote.subarray(38, 77).toString()).toBe('application/vnd.oasis.opendocument.text')
    expect(entradas(pacote)).toEqual([
      'mimetype',
      'content.xml',
      'styles.xml',
      'meta.xml',
      'META-INF/manifest.xml',
    ])
    const manifesto = await entryOf(odt, 'META-INF/manifest.xml')
    for (const parte of ['content.xml', 'styles.xml', 'meta.xml']) {
      expect(manifesto).toContain(`manifest:full-path="${parte}"`)
    }
    const conteudo = await entryOf(odt, 'content.xml')
    expect(conteudo).toContain('Ata da reunião de terça.')
    expect(conteudo).toMatch(/<text:note text:id="nota-rodape-1" text:note-class="footnote">/)
    expect(conteudo).toContain('Fonte: ata anterior.')
    expect(await entryOf(odt, 'styles.xml')).toContain('<office:master-styles>')

    await expect(session.window).toHaveTitle(/^ata\.docx/)
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })
})
