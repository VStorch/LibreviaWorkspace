/// <reference lib="dom" />
// The `evaluate` body runs in the renderer, but is compiled in Node's scope.

import { existsSync } from 'node:fs'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'

/**
 * Anchored objects drawn at their sheet position: in Word they do not push the text and may sit
 * behind it. The corpus manual template has two positioned boxes, a rotated image and an object
 * behind the text. The corpus is not in the repository: without `LIBREVIA_CORPUS_DOC`, the tests
 * are skipped.
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
    // The position depends on the sheet the anchor paragraph fell on, and only exists after
    // measuring.
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
    // Rotated, the mark is a narrow band on the left, before the text column.
    expect(caixa!.largura).toBeLessThan(caixa!.larguraFolha / 2)
    expect(caixa!.esquerda).toBeLessThan(caixa!.larguraFolha / 4)
  })

  test('o objeto de trás fica atrás do texto', async () => {
    // `behindDoc` is cover decoration and watermark: on top, it would cover the text.
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
    // Drawings anchored in the band, rotated a quarter turn: 28.6 mm standing in a 10 mm band.
    const giradas = await session.window.evaluate(
      () =>
        Array.from(document.querySelectorAll('.paper-float'))
          .map((node) => getComputedStyle(node as HTMLElement).transform)
          .filter((transform) => transform !== 'none').length,
    )

    // The body's mark, the header's and the footer's, on every sheet.
    expect(giradas).toBeGreaterThanOrEqual(3)
  })

  test('as marcas da faixa repetem em toda folha', async () => {
    // The band repeats, and so does what is anchored in it: it belongs to the page.
    const porFolha = await session.window.evaluate(() =>
      Array.from(document.querySelectorAll('.paper-bands')).map(
        (banda) => banda.querySelectorAll('.paper-float').length,
      ),
    )

    expect(porFolha.length).toBeGreaterThan(1)
    expect(Math.min(...porFolha)).toBeGreaterThanOrEqual(3)
  })

  test('os objetos não ocupam lugar no fluxo do texto', async () => {
    // The paragraph anchoring the 286 mm mark keeps its line height.
    const alturas = await session.window.evaluate(() =>
      Array.from(document.querySelectorAll('.page__content > *')).map(
        (node) => (node as HTMLElement).offsetHeight,
      ),
    )

    expect(Math.max(...alturas)).toBeLessThan(200)
  })
})
