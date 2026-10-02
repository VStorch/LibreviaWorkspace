import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithFootnote } from './fixtures.js'

/**
 * Arquivo → Exportar como → HTML… e Markdown… (M11). Exportar escreve um arquivo
 * novo: o documento continua no caminho dele e sem alteração pendente.
 */
test.describe('exportar como HTML e Markdown', () => {
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

    // O documento continua sendo o .docx, limpo, e o arquivo de origem intacto.
    await expect(session.window).toHaveTitle(/^ata\.docx/)
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect((await stat(origem)).mtimeMs).toBe(antes)
  })
})
