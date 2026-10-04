/// <reference lib="dom" />
// O corpo de `evaluate` roda no renderer, mas é compilado no escopo do Node.

import { existsSync } from 'node:fs'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'

/**
 * Objetos ancorados desenhados na posição da folha: no Word não empurram o texto e
 * podem ficar atrás dele. O modelo de manual do corpus traz duas caixas
 * posicionadas, uma imagem girada e um objeto atrás do texto. O corpus não entra no
 * repositório: sem `LIBREVIA_CORPUS_DOC`, os testes são pulados.
 */
const MODELO = process.env['LIBREVIA_CORPUS_DOC'] ?? ''

test.describe('objetos ancorados', () => {
  let session: Session

  test.skip(
    MODELO === '' || !existsSync(MODELO),
    'requer LIBREVIA_CORPUS_DOC apontando para o documento do corpus',
  )

  test.beforeEach(async () => {
    session = await launch()
    await stubDialogs(session.app, { open: MODELO, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()
    // A posição depende da folha em que o parágrafo âncora caiu, e só existe depois da medição.
    await expect(session.window.locator('.paper-float').first()).toBeVisible()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('o texto das caixas aparece na posição delas, e não emendado na linha', async () => {
    const caixas = session.window.locator('.paper-float--text')
    await expect(caixas.filter({ hasText: 'Título' })).toHaveCount(1)
    await expect(caixas.filter({ hasText: 'Subtitulo' })).toHaveCount(1)
  })

  test('a marca girada fica na lateral, à esquerda da coluna de texto', async () => {
    const caixa = await session.window.evaluate(() => {
      const paper = document.querySelector('.paper') as HTMLElement
      const marca = document.querySelector('.paper-floats--behind img') as HTMLElement | null
      if (marca === null) return null

      const folha = paper.getBoundingClientRect()
      const box = marca.getBoundingClientRect()
      return { esquerda: box.left - folha.left, largura: box.width, larguraFolha: folha.width }
    })

    expect(caixa).not.toBeNull()
    // Girada, a marca é uma faixa estreita à esquerda, antes da coluna de texto.
    expect(caixa!.largura).toBeLessThan(caixa!.larguraFolha / 2)
    expect(caixa!.esquerda).toBeLessThan(caixa!.larguraFolha / 4)
  })

  test('o objeto de trás fica atrás do texto', async () => {
    // `behindDoc` é decoração de capa e marca d'água: por cima, cobriria o texto.
    const ordem = await session.window.evaluate(() => {
      const atras = document.querySelector('.paper-floats--behind') as HTMLElement | null
      const coluna = document.querySelector('.pages__column') as HTMLElement | null
      if (atras === null || coluna === null) return null
      return {
        atras: Number(getComputedStyle(atras).zIndex),
        texto: Number(getComputedStyle(coluna).zIndex),
      }
    })

    expect(ordem).not.toBeNull()
    expect(ordem!.atras).toBeLessThan(ordem!.texto)
  })

  test('as marcas do cabeçalho e do rodapé giram como no arquivo', async () => {
    // Desenhos ancorados na faixa, girados um quarto de volta: 28,6 mm em pé numa banda de 10 mm.
    const giradas = await session.window.evaluate(
      () =>
        Array.from(document.querySelectorAll('.paper-float'))
          .map((node) => getComputedStyle(node as HTMLElement).transform)
          .filter((transform) => transform !== 'none').length,
    )

    // A marca do corpo, a do cabeçalho e a do rodapé — em cada folha.
    expect(giradas).toBeGreaterThanOrEqual(3)
  })

  test('as marcas da faixa repetem em toda folha', async () => {
    // A faixa repete, e o que está ancorado nela também: pertence à página.
    const porFolha = await session.window.evaluate(() =>
      Array.from(document.querySelectorAll('.paper-bands')).map(
        (banda) => banda.querySelectorAll('.paper-float').length,
      ),
    )

    expect(porFolha.length).toBeGreaterThan(1)
    expect(Math.min(...porFolha)).toBeGreaterThanOrEqual(3)
  })

  test('os objetos não ocupam lugar no fluxo do texto', async () => {
    // O parágrafo que ancora a marca de 286 mm continua com altura de linha.
    const alturas = await session.window.evaluate(() =>
      Array.from(document.querySelectorAll('.page__content > *')).map(
        (node) => (node as HTMLElement).offsetHeight,
      ),
    )

    expect(Math.max(...alturas)).toBeLessThan(200)
  })
})
