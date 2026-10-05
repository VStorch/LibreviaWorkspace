/// <reference lib="dom" />
// The `evaluate` body runs in the renderer, but is compiled in Node's scope.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithAnchoredScreenshot, docxWithIndentedScreenshot, docxWithStretchedImage } from './fixtures.js'

/**
 * An image comes out at the size the document asked for: `wp:extent` gives the size on the page,
 * without the file's ratio, and the box must not measure zero until the bytes decode, which is when
 * pagination measures.
 */
test.describe('imagens do documento', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-imagens-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('a imagem esticada continua esticada, e não volta ao quadrado', async () => {
    const origem = join(pasta, 'imagem.docx')
    await writeFile(origem, await docxWithStretchedImage())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })

    await menu(session, 'open')
    // `img[src]`: ProseMirror puts an empty <img> next to every atomic node.
    const imagem = session.window.locator('.page__content img[src^="data:"]')
    await expect(imagem).toBeVisible()

    const caixa = await imagem.evaluate((node) => {
      const img = node as HTMLImageElement
      const box = img.getBoundingClientRect()
      return { largura: box.width, altura: box.height, natural: img.naturalWidth === img.naturalHeight }
    })

    // The file is square; the document asks for four by one.
    expect(caixa.natural).toBe(true)
    expect(caixa.largura / caixa.altura).toBeCloseTo(4, 1)
    expect(caixa.largura).toBeCloseTo(400, 0)
  })

  test('a imagem ancorada no lugar do parágrafo ocupa altura no texto', async () => {
    // LibreOffice writes a screenshot anchored to the paragraph, without offset and centered: it
    // takes height in the flow.
    const origem = join(pasta, 'captura.docx')
    await writeFile(origem, await docxWithAnchoredScreenshot())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })

    await menu(session, 'open')
    const imagem = session.window.locator('.page__content img[src^="data:"]')
    await expect(imagem).toBeVisible()

    const medidas = await session.window.evaluate(() => {
      const img = document.querySelector('.page__content img[src^="data:"]') as HTMLImageElement
      const depois = Array.from(document.querySelectorAll('.page__content > *')).find((node) =>
        (node.textContent ?? '').startsWith('Depois'),
      ) as HTMLElement | null

      return {
        flutuantes: document.querySelectorAll('.paper-float').length,
        fim: img.getBoundingClientRect().bottom,
        seguinte: depois?.getBoundingClientRect().top ?? 0,
      }
    })

    // No positioned copy, and the next paragraph starts after it.
    expect(medidas.flutuantes).toBe(0)
    expect(medidas.seguinte).toBeGreaterThanOrEqual(medidas.fim - 1)
  })

  test('o parágrafo da captura ocupa a altura dela mais uma linha', async () => {
    // The image is a block, not a word: inline, the font descender would be left over, which Word
    // does not charge. Word does charge the paragraph's empty line, because the frame takes the
    // column and the line moves down; LibreOffice leaves one line height between two screenshots.
    const origem = join(pasta, 'altura.docx')
    await writeFile(origem, await docxWithAnchoredScreenshot())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })

    await menu(session, 'open')
    const imagem = session.window.locator('.page__content img[src^="data:"]')
    await expect(imagem).toBeVisible()

    const medidas = await session.window.evaluate(() => {
      const img = document.querySelector('.page__content img[src^="data:"]') as HTMLImageElement
      // The paragraph, not the direct parent, which is the NodeView handle frame.
      const bloco = img.closest('p') as HTMLElement
      return {
        sobra: bloco.getBoundingClientRect().height - img.getBoundingClientRect().height,
        entrelinha: Number.parseFloat(getComputedStyle(bloco).lineHeight),
      }
    })

    // One line, not the font descender: what is left over is the paragraph's line height.
    expect(medidas.sobra).toBeCloseTo(medidas.entrelinha, 0)
  })

  test('a captura ocupa a coluna mesmo dentro de um parágrafo recuado', async () => {
    // In Word an anchored screenshot positions itself by the column, and the indent does not narrow
    // it.
    const origem = join(pasta, 'recuo.docx')
    await writeFile(origem, await docxWithIndentedScreenshot())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })

    await menu(session, 'open')
    const imagem = session.window.locator('.page__content img[src^="data:"]')
    await expect(imagem).toBeVisible()

    const medidas = await session.window.evaluate(() => {
      const img = document.querySelector('.page__content img[src^="data:"]') as HTMLImageElement
      const legenda = document.querySelector('.page__content p') as HTMLElement
      return {
        imagem: img.getBoundingClientRect().width,
        // The NodeView puts the measure in the style, to beat the sheet's `height: auto`.
        pedida: Number.parseFloat(img.style.width),
        // The indent is padding: the paragraph box is still the column.
        recuoDaLegenda: Number.parseFloat(getComputedStyle(legenda).paddingLeft),
      }
    })

    // 3810000 EMU is 400 px, the width the file asks for.
    expect(medidas.imagem).toBeCloseTo(medidas.pedida, 0)

    // And the indent still exists for the text: half an inch is 48 px.
    expect(medidas.recuoDaLegenda).toBeCloseTo(48, 0)
  })
})
