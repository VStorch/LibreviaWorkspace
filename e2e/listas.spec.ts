import { execFile } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithMultilevelList, entryOf } from './fixtures.js'

/**
 * Listas multinível: a marca de cada item é a que o Word desenharia.
 *
 * O contador do CSS recomeça a cada lista e só conhece um formato; o documento
 * de teste tem o que ele não sabe fazer — o segundo nível compondo o primeiro
 * (`1.a)`), a lista que continua do outro lado de um parágrafo (`3.`) e um
 * reinício no mesmo documento (`10.`).
 */
test.describe('listas multinível', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-listas-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  const marcas = (session: Session) =>
    session.window
      .locator('.ProseMirror li[data-label]')
      .evaluateAll((items) => items.map((item) => item.getAttribute('data-label')))

  async function abrirMultinivel(): Promise<string> {
    const origem = join(pasta, 'entrada.docx')
    await writeFile(origem, await docxWithMultilevelList())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Dez')
    return origem
  }

  test('a tela numera como o Word: níveis compostos, continuação e reinício', async () => {
    await abrirMultinivel()
    await expect.poll(() => marcas(session)).toEqual(['1.', '1.a)', '1.b)', '2.', '3.', '10.'])

    // E é a marca que se vê: o `::before` do parágrafo desenha a do item.
    const desenhada = await session.window
      .locator('.ProseMirror li[data-label] > p')
      .nth(1)
      .evaluate((paragrafo) => getComputedStyle(paragrafo, '::before').content)
    expect(desenhada).toBe('"1.a)"')
  })

  test('o PDF mostra as mesmas marcas da tela', async () => {
    test.skip(!(await temPdftotext()), 'pdftotext não instalado')
    await abrirMultinivel()
    const destino = join(pasta, 'saida.pdf')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'export-pdf')

    await expect.poll(() => textoDoPdf(destino), { timeout: 30_000 }).toContain('1.a)')
    const texto = await textoDoPdf(destino)
    expect(texto).toMatch(/1\.b\)\s+Um-b/)
    expect(texto).toMatch(/3\.\s+Três/)
    expect(texto).toMatch(/10\.\s+Dez/)
  })

  test('abrir e salvar sem editar devolve o numbering.xml byte a byte', async () => {
    const origem = await abrirMultinivel()
    const destino = join(pasta, 'saida.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    expect(await entryOf(destino, 'word/numbering.xml')).toBe(await entryOf(origem, 'word/numbering.xml'))
  })

  test('Tab desce o item um nível e Shift+Tab o devolve', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('1. Primeiro')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Segundo')
    await expect.poll(() => marcas(session)).toEqual(['1.', '2.'])

    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('Tab')
    await expect.poll(() => marcas(session)).toEqual(['1.', 'a.'])

    await session.window.keyboard.press('Shift+Tab')
    await expect.poll(() => marcas(session)).toEqual(['1.', '2.'])
  })
})

async function temPdftotext(): Promise<boolean> {
  try {
    await promisify(execFile)('pdftotext', ['-v'])
    return true
  } catch {
    return false
  }
}

async function textoDoPdf(caminho: string): Promise<string> {
  try {
    const { stdout } = await promisify(execFile)('pdftotext', ['-layout', caminho, '-'])
    return stdout
  } catch {
    return ''
  }
}
