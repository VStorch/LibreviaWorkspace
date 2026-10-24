import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithFootnote, entryOf } from './fixtures.js'

/**
 * File → Export as → HTML, Markdown and ODT write a new file: the document stays at its path and
 * without pending changes. Word is a save as, offering only `.docx`.
 */
test.describe('export as Word, HTML, Markdown and ODT', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-export-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  const read = (path: string) => () => readFile(path, 'utf8').catch(() => '')

  /** The ZIP entry names, from the central directory, in file order. */
  const entries = (zip: Buffer): string[] => {
    const end = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
    const names: string[] = []
    let i = zip.readUInt32LE(end + 16)
    for (let n = zip.readUInt16LE(end + 10); n > 0; n--) {
      const name = zip.readUInt16LE(i + 28)
      names.push(zip.subarray(i + 46, i + 46 + name).toString('utf8'))
      i += 46 + name + zip.readUInt16LE(i + 30) + zip.readUInt16LE(i + 32)
    }
    return names
  }

  test('exports the open document without changing its path or marking it modified', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithFootnote())
    const before = (await stat(sourceFile)).mtimeMs
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Ata da reunião')

    const html = join(directory, 'ata.html')
    await stubDialogs(session.app, { save: html })
    await menu(session, 'export-html')
    await expect.poll(read(html), { timeout: 15_000 }).toContain('</html>')
    const page = await readFile(html, 'utf8')
    expect(page).toContain('<html lang="pt-BR">')
    expect(page).toContain('<title>ata</title>')
    expect(page).toContain('<p>Ata da reunião de terça.</p>')
    expect(page).toContain('<a href="#footnote-1" id="ref-footnote-1">1</a>')
    expect(page).toContain('Fonte: ata anterior.')
    expect(page).not.toMatch(/<script/i)

    const markdown = join(directory, 'ata.md')
    await stubDialogs(session.app, { save: markdown })
    await menu(session, 'export-markdown')
    await expect.poll(read(markdown), { timeout: 15_000 }).toContain('[^1]:')
    const text = await readFile(markdown, 'utf8')
    expect(text).toContain('Ata da reunião de terça.\n\nSegundo parágrafo, com uma nota.[^1]')
    expect(text).toContain('[^1]: Fonte: ata anterior.')

    // The document is still the .docx, clean, and the source file intact.
    await expect(session.window).toHaveTitle(/^ata\.docx/)
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect((await stat(sourceFile)).mtimeMs).toBe(before)
  })

  test('exports an OpenDocument package with its parts to ODT, without marking the document modified', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithFootnote())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Ata da reunião')

    const odt = join(directory, 'ata.odt')
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
    const archive = await readFile(odt)
    expect(archive.subarray(30, 38).toString()).toBe('mimetype')
    expect(archive.subarray(38, 77).toString()).toBe('application/vnd.oasis.opendocument.text')
    expect(entries(archive)).toEqual([
      'mimetype',
      'content.xml',
      'styles.xml',
      'meta.xml',
      'META-INF/manifest.xml',
    ])
    const manifest = await entryOf(odt, 'META-INF/manifest.xml')
    for (const part of ['content.xml', 'styles.xml', 'meta.xml']) {
      expect(manifest).toContain(`manifest:full-path="${part}"`)
    }
    const contents = await entryOf(odt, 'content.xml')
    expect(contents).toContain('Ata da reunião de terça.')
    expect(contents).toMatch(/<text:note text:id="footnote-1" text:note-class="footnote">/)
    expect(contents).toContain('Fonte: ata anterior.')
    expect(await entryOf(odt, 'styles.xml')).toContain('<office:master-styles>')

    await expect(session.window).toHaveTitle(/^ata\.docx/)
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })

  test('Word saves the document as .docx, offering only that format, and goes on from there', async () => {
    await menu(session, 'new-document')
    const text = session.window.locator('.pages__column .ProseMirror')
    await text.click()
    await session.window.keyboard.type('Relatório mensal')

    const typed = join(directory, 'relatorio')
    await session.app.evaluate(({ dialog }, path) => {
      const offered = globalThis as { offered?: unknown }
      dialog.showSaveDialog = (async (...args: unknown[]) => {
        offered.offered = args.at(-1)
        return { canceled: false, filePath: path }
      }) as typeof dialog.showSaveDialog
    }, typed)
    await menu(session, 'export-word')

    const docx = `${typed}.docx`
    await expect
      .poll(() => entryOf(docx, 'word/document.xml').catch(() => ''), { timeout: 15_000 })
      .toContain('Relatório mensal')
    const offered = await session.app.evaluate(
      () =>
        (globalThis as { offered?: { defaultPath: string; filters: { extensions: string[] }[] } }).offered,
    )
    expect(offered?.defaultPath).toMatch(/\.docx$/)
    expect(offered?.filters.map((filter) => filter.extensions)).toEqual([['docx']])
    await expect(session.window).toHaveTitle(/^relatorio\.docx/)
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })
})
