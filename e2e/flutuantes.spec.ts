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
const TEMPLATE = process.env['LIBREVIA_CORPUS_DOC'] ?? ''

test.describe('anchored objects', () => {
  let session: Session

  test.skip(
    TEMPLATE === '' || !existsSync(TEMPLATE),
    'requer LIBREVIA_CORPUS_DOC apontando para o documento do corpus',
  )

  test.beforeEach(async () => {
    session = await launch()
    await stubDialogs(session.app, { open: TEMPLATE, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()
    // The position depends on the sheet the anchor paragraph fell on, and only exists after
    // measuring.
    await expect(session.window.locator('.paper-float').first()).toBeVisible()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('the text of the boxes appears at their position, not joined into the line', async () => {
    const boxes = session.window.locator('.paper-float--text')
    await expect(boxes.filter({ hasText: 'Título' })).toHaveCount(1)
    await expect(boxes.filter({ hasText: 'Subtitulo' })).toHaveCount(1)
  })

  test('the rotated mark stays at the side, left of the text column', async () => {
    const boxRect = await session.window.evaluate(() => {
      const paper = document.querySelector('.paper') as HTMLElement
      const mark = document.querySelector('.paper-floats--behind img') as HTMLElement | null
      if (mark === null) return null

      const sheet = paper.getBoundingClientRect()
      const box = mark.getBoundingClientRect()
      return { left: box.left - sheet.left, width: box.width, sheetWidth: sheet.width }
    })

    expect(boxRect).not.toBeNull()
    // Rotated, the mark is a narrow band on the left, before the text column.
    expect(boxRect!.width).toBeLessThan(boxRect!.sheetWidth / 2)
    expect(boxRect!.left).toBeLessThan(boxRect!.sheetWidth / 4)
  })

  test('the object behind stays behind the text', async () => {
    // `behindDoc` is cover decoration and watermark: on top, it would cover the text.
    const order = await session.window.evaluate(() => {
      const behind = document.querySelector('.paper-floats--behind') as HTMLElement | null
      const column = document.querySelector('.pages__column') as HTMLElement | null
      if (behind === null || column === null) return null
      return {
        behind: Number(getComputedStyle(behind).zIndex),
        text: Number(getComputedStyle(column).zIndex),
      }
    })

    expect(order).not.toBeNull()
    expect(order!.behind).toBeLessThan(order!.text)
  })

  test('the header and footer marks rotate as in the file', async () => {
    // Drawings anchored in the band, rotated a quarter turn: 28.6 mm standing in a 10 mm band.
    const rotated = await session.window.evaluate(
      () =>
        Array.from(document.querySelectorAll('.paper-float'))
          .map((node) => getComputedStyle(node as HTMLElement).transform)
          .filter((transform) => transform !== 'none').length,
    )

    // The body's mark, the header's and the footer's, on every sheet.
    expect(rotated).toBeGreaterThanOrEqual(3)
  })

  test('the band marks repeat on every sheet', async () => {
    // The band repeats, and so does what is anchored in it: it belongs to the page.
    const perSheet = await session.window.evaluate(() =>
      Array.from(document.querySelectorAll('.paper-bands')).map(
        (band) => band.querySelectorAll('.paper-float').length,
      ),
    )

    expect(perSheet.length).toBeGreaterThan(1)
    expect(Math.min(...perSheet)).toBeGreaterThanOrEqual(3)
  })

  test('the objects take no place in the text flow', async () => {
    // The paragraph anchoring the 286 mm mark keeps its line height.
    const heights = await session.window.evaluate(() =>
      Array.from(document.querySelectorAll('.page__content > *')).map(
        (node) => (node as HTMLElement).offsetHeight,
      ),
    )

    expect(Math.max(...heights)).toBeLessThan(200)
  })
})
