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
test.describe('document images', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-images-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  test('a stretched image stays stretched and does not go back to square', async () => {
    const sourceFile = join(directory, 'imagem.docx')
    await writeFile(sourceFile, await docxWithStretchedImage())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })

    await menu(session, 'open')
    // `img[src]`: ProseMirror puts an empty <img> next to every atomic node.
    const image = session.window.locator('.page__content img[src^="data:"]')
    await expect(image).toBeVisible()

    const boxRect = await image.evaluate((node) => {
      const img = node as HTMLImageElement
      const box = img.getBoundingClientRect()
      return { width: box.width, height: box.height, natural: img.naturalWidth === img.naturalHeight }
    })

    // The file is square; the document asks for four by one.
    expect(boxRect.natural).toBe(true)
    expect(boxRect.width / boxRect.height).toBeCloseTo(4, 1)
    expect(boxRect.width).toBeCloseTo(400, 0)
  })

  test("an image anchored in the paragraph's place takes up height in the text", async () => {
    // LibreOffice writes a screenshot anchored to the paragraph, without offset and centered: it
    // takes height in the flow.
    const sourceFile = join(directory, 'captura.docx')
    await writeFile(sourceFile, await docxWithAnchoredScreenshot())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })

    await menu(session, 'open')
    const image = session.window.locator('.page__content img[src^="data:"]')
    await expect(image).toBeVisible()

    const measures = await session.window.evaluate(() => {
      const img = document.querySelector('.page__content img[src^="data:"]') as HTMLImageElement
      const after = Array.from(document.querySelectorAll('.page__content > *')).find((node) =>
        (node.textContent ?? '').startsWith('Depois'),
      ) as HTMLElement | null

      return {
        floating: document.querySelectorAll('.paper-float').length,
        end: img.getBoundingClientRect().bottom,
        next: after?.getBoundingClientRect().top ?? 0,
      }
    })

    // No positioned copy, and the next paragraph starts after it.
    expect(measures.floating).toBe(0)
    expect(measures.next).toBeGreaterThanOrEqual(measures.end - 1)
  })

  test('the screenshot paragraph takes its height plus one line', async () => {
    // The image is a block, not a word: inline, the font descender would be left over, which Word
    // does not charge. Word does charge the paragraph's empty line, because the frame takes the
    // column and the line moves down; LibreOffice leaves one line height between two screenshots.
    const sourceFile = join(directory, 'altura.docx')
    await writeFile(sourceFile, await docxWithAnchoredScreenshot())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })

    await menu(session, 'open')
    const image = session.window.locator('.page__content img[src^="data:"]')
    await expect(image).toBeVisible()

    const measures = await session.window.evaluate(() => {
      const img = document.querySelector('.page__content img[src^="data:"]') as HTMLImageElement
      // The paragraph, not the direct parent, which is the NodeView handle frame.
      const block = img.closest('p') as HTMLElement
      return {
        leftover: block.getBoundingClientRect().height - img.getBoundingClientRect().height,
        lineSpacing: Number.parseFloat(getComputedStyle(block).lineHeight),
      }
    })

    // One line, not the font descender: what is left over is the paragraph's line height.
    expect(measures.leftover).toBeCloseTo(measures.lineSpacing, 0)
  })

  test('the screenshot fills the column even inside an indented paragraph', async () => {
    // In Word an anchored screenshot positions itself by the column, and the indent does not narrow
    // it.
    const sourceFile = join(directory, 'recuo.docx')
    await writeFile(sourceFile, await docxWithIndentedScreenshot())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })

    await menu(session, 'open')
    const image = session.window.locator('.page__content img[src^="data:"]')
    await expect(image).toBeVisible()

    const measures = await session.window.evaluate(() => {
      const img = document.querySelector('.page__content img[src^="data:"]') as HTMLImageElement
      const caption = document.querySelector('.page__content p') as HTMLElement
      return {
        image: img.getBoundingClientRect().width,
        // The NodeView puts the measure in the style, to beat the sheet's `height: auto`.
        requested: Number.parseFloat(img.style.width),
        // The indent is padding: the paragraph box is still the column.
        captionIndent: Number.parseFloat(getComputedStyle(caption).paddingLeft),
      }
    })

    // 3810000 EMU is 400 px, the width the file asks for.
    expect(measures.image).toBeCloseTo(measures.requested, 0)

    // And the indent still exists for the text: half an inch is 48 px.
    expect(measures.captionIndent).toBeCloseTo(48, 0)
  })
})
