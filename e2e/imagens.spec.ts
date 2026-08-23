/// <reference lib="dom" />
// O corpo de `evaluate` roda no renderer, mas é compilado no escopo do Node.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithStretchedImage } from './fixtures.js'

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
})
