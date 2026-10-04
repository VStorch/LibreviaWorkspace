/// <reference lib="dom" />
// O corpo de `evaluate` roda no renderer, mas é compilado no escopo do Node.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithAnchoredScreenshot, docxWithIndentedScreenshot, docxWithStretchedImage } from './fixtures.js'

/**
 * A imagem sai do tamanho que o documento pediu: `wp:extent` dá o tamanho na
 * página, sem a proporção do arquivo, e a caixa não pode medir zero até os bytes
 * decodificarem, que é quando a paginação mede.
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
    // `img[src]`: o ProseMirror põe um <img> vazio ao lado de todo nó atômico.
    const imagem = session.window.locator('.page__content img[src^="data:"]')
    await expect(imagem).toBeVisible()

    const caixa = await imagem.evaluate((node) => {
      const img = node as HTMLImageElement
      const box = img.getBoundingClientRect()
      return { largura: box.width, altura: box.height, natural: img.naturalWidth === img.naturalHeight }
    })

    // O arquivo é quadrado; o documento pede quatro por um.
    expect(caixa.natural).toBe(true)
    expect(caixa.largura / caixa.altura).toBeCloseTo(4, 1)
    expect(caixa.largura).toBeCloseTo(400, 0)
  })

  test('a imagem ancorada no lugar do parágrafo ocupa altura no texto', async () => {
    // O LibreOffice grava a captura de tela ancorada ao parágrafo, sem deslocamento e
    // centralizada: ela ocupa altura no fluxo.
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

    // Nenhuma cópia posicionada, e o parágrafo seguinte começa depois dela.
    expect(medidas.flutuantes).toBe(0)
    expect(medidas.seguinte).toBeGreaterThanOrEqual(medidas.fim - 1)
  })

  test('o parágrafo da captura ocupa a altura dela mais uma linha', async () => {
    // A imagem é bloco, e não palavra: inline sobraria a descida da fonte, que o Word
    // não cobra. A linha vazia do parágrafo o Word cobra, porque o quadro ocupa a
    // coluna e ela desce; o LibreOffice deixa uma entrelinha entre duas capturas.
    const origem = join(pasta, 'altura.docx')
    await writeFile(origem, await docxWithAnchoredScreenshot())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })

    await menu(session, 'open')
    const imagem = session.window.locator('.page__content img[src^="data:"]')
    await expect(imagem).toBeVisible()

    const medidas = await session.window.evaluate(() => {
      const img = document.querySelector('.page__content img[src^="data:"]') as HTMLImageElement
      // O parágrafo, e não o pai direto, que é a moldura das alças do NodeView.
      const bloco = img.closest('p') as HTMLElement
      return {
        sobra: bloco.getBoundingClientRect().height - img.getBoundingClientRect().height,
        entrelinha: Number.parseFloat(getComputedStyle(bloco).lineHeight),
      }
    })

    // Uma linha, e não a descida da fonte: a sobra é a entrelinha do parágrafo.
    expect(medidas.sobra).toBeCloseTo(medidas.entrelinha, 0)
  })

  test('a captura ocupa a coluna mesmo dentro de um parágrafo recuado', async () => {
    // No Word a captura ancorada se posiciona pela coluna, e o recuo não a estreita.
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
        // O NodeView põe a medida no estilo, para vencer o `height: auto` da folha.
        pedida: Number.parseFloat(img.style.width),
        // O recuo é preenchimento: a caixa do parágrafo continua sendo a coluna.
        recuoDaLegenda: Number.parseFloat(getComputedStyle(legenda).paddingLeft),
      }
    })

    // 3810000 EMU são 400 px, a largura que o arquivo pede.
    expect(medidas.imagem).toBeCloseTo(medidas.pedida, 0)

    // E o recuo continua existindo para o texto: meia polegada são 48 px.
    expect(medidas.recuoDaLegenda).toBeCloseTo(48, 0)
  })
})
