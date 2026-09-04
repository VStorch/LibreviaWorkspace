/// <reference lib="dom" />
// O corpo de `evaluate` roda no renderer, mas é compilado no escopo do Node.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import {
  docxWithAnchoredScreenshot,
  docxWithIndentedScreenshot,
  docxWithStretchedImage,
} from './fixtures.js'

/**
 * A imagem sai do tamanho que o documento pediu.
 *
 * `wp:extent` diz de que tamanho a imagem é **na página**, e esse tamanho não
 * precisa ter a proporção do arquivo. Só a largura chegava até aqui: a altura
 * ficava por conta do navegador, que a tira da proporção natural do arquivo —
 * então imagem esticada de propósito voltava ao quadrado, e até os bytes
 * decodificarem a caixa media zero, bem no momento em que a paginação mede a
 * folha.
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
    // `img[src]` de verdade: o ProseMirror põe um <img> vazio de separação ao
    // lado de todo nó atômico, e ele também casaria com um seletor solto.
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
    // `wp:anchor` não quer dizer "fora do fluxo": é assim que o LibreOffice
    // grava captura de tela — ancorada ao parágrafo, sem deslocamento,
    // centralizada na coluna. Tratada como posição na folha, ela deixava de
    // ocupar altura: o texto se fechava por cima dela, e um documento de trinta
    // capturas encolhia de quinze folhas para quatro.
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
    // Duas coisas de uma vez. A imagem é bloco e não palavra: inline ela
    // repousaria sobre a linha de base e sobraria por baixo a descida da fonte,
    // que o Word não cobra. E o parágrafo dela tem uma linha vazia, que o Word
    // cobra: a captura ancorada é um quadro que flutua, e com ela ocupando a
    // coluna inteira a linha não cabe ao lado e vai para baixo.
    //
    // Medido no LibreOffice: entre duas capturas encostadas ele deixa
    // exatamente uma entrelinha.
    const origem = join(pasta, 'altura.docx')
    await writeFile(origem, await docxWithAnchoredScreenshot())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })

    await menu(session, 'open')
    const imagem = session.window.locator('.page__content img[src^="data:"]')
    await expect(imagem).toBeVisible()

    const medidas = await session.window.evaluate(() => {
      const img = document.querySelector('.page__content img[src^="data:"]') as HTMLImageElement
      const bloco = img.parentElement as HTMLElement
      return {
        sobra: bloco.getBoundingClientRect().height - img.getBoundingClientRect().height,
        entrelinha: Number.parseFloat(getComputedStyle(bloco).lineHeight),
      }
    })

    // Uma linha, e não a descida da fonte: a sobra é a entrelinha do parágrafo.
    expect(medidas.sobra).toBeCloseTo(medidas.entrelinha, 0)
  })

  test('a captura ocupa a coluna mesmo dentro de um parágrafo recuado', async () => {
    // O recuo do Word é uma medida, e a captura ancorada não é texto: no Word
    // ela se posiciona pela coluna. Espremida pelo recuo, ela encolhia também
    // em altura — e a legenda seguinte passava a caber numa folha em que o
    // LibreOffice já não a punha.
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
        pedida: Number(img.getAttribute('width')),
        // O recuo é preenchimento e não margem: a caixa do parágrafo continua
        // sendo a coluna, e é por dentro dela que o texto anda.
        recuoDaLegenda: Number.parseFloat(getComputedStyle(legenda).paddingLeft),
      }
    })

    // 3810000 EMU são 400 px: a largura que o arquivo pede, e não a caixa do
    // parágrafo recuado.
    expect(medidas.imagem).toBeCloseTo(medidas.pedida, 0)

    // E o recuo continua existindo para o texto: meia polegada são 48 px.
    expect(medidas.recuoDaLegenda).toBeCloseTo(48, 0)
  })
})
