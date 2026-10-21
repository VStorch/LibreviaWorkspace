/// <reference lib="dom" />
// The `evaluate` body runs in the renderer, but is compiled in Node's scope.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithHeaderGrid, entryOf } from './fixtures.js'

/**
 * A grid header: the corporate one is a table, with the logo in a cell merged across rows, the
 * title beside it and the numbering on the right.
 */
test.describe('grid header', () => {
  let session: Session
  let directory: string
  let destinationFile: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-header-'))
    destinationFile = join(directory, 'salva.docx')
    session = await launch()

    const sourceFile = join(directory, 'grade.docx')
    await writeFile(sourceFile, await docxWithHeaderGrid())
    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  test('the header table is drawn as a table', async () => {
    const grid = session.window.locator('.band--header .band__grid').first()
    await expect(grid).toBeVisible()
    await expect(grid.locator('tr')).toHaveCount(4)
    await expect(grid.locator('td')).toHaveCount(5)
  })

  test('the merged cell grows instead of leaving the row empty', async () => {
    // In OOXML vertical merging is not height: the top cell says `restart` and the lower one shows
    // as an empty cell. Drawn as a real cell, it would open a blank band under the logo.
    const stamp = session.window.locator('.band--header .band__grid td', { hasText: 'Selo' }).first()
    await expect(stamp).toHaveAttribute('rowspan', '4')
  })

  test('the body moves down below the header, without meeting it', async () => {
    // The top margin is a floor: with a taller header, Word and LibreOffice push the body down.
    const measures = await session.window.evaluate(() => {
      const band = document.querySelector('.band--header') as HTMLElement | null
      const first = document.querySelector('.page__content > *') as HTMLElement | null
      if (band === null || first === null) return null
      return {
        bandEnd: band.getBoundingClientRect().bottom,
        start: first.getBoundingClientRect().top,
      }
    })

    expect(measures).not.toBeNull()
    expect(measures!.start).toBeGreaterThanOrEqual(measures!.bandEnd)
  })

  test('the header text is typeable and returns to the file', async () => {
    // The header title is editable, and only that piece's `w:t` is rewritten: table, borders and
    // merging stay byte for byte.
    const title = session.window
      .locator('.band--header .band__text')
      .filter({ hasText: 'Título do documento' })
      .first()

    await title.click()
    await session.window.keyboard.press('ControlOrMeta+a')
    await session.window.keyboard.type('Título corrigido')

    // Text goes out on `blur`: changing the setup redraws the sheets, and redrawing under someone
    // typing would take the cursor away.
    await session.window.locator('.ProseMirror').click()
    await expect(session.window.locator('.band--header .band__grid')).toContainText('Título corrigido')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const header = await entryOf(destinationFile, 'word/header1.xml')
    expect(header).toContain('Título corrigido')
    expect(header).not.toContain('Título do documento')

    // The rest of the part was not regenerated: the grid is still there, with the merge this writer
    // could not produce from scratch.
    expect(header).toContain('w:vMerge w:val="restart"')
    expect(header).toContain('Chamado 10001')
  })

  test('what has no text of its own in the file does not receive the cursor', async () => {
    // The margin around the band still belongs to the body: clicking it must not take the cursor
    // out of the text.
    const editable = await session.window.locator('.band--header .band__text').count()
    const pieces = await session.window.locator('.band--header .band__grid td').count()

    expect(editable).toBeGreaterThan(0)
    expect(editable).toBeLessThan(pieces + 1)
    await expect(session.window.locator('.band--header')).not.toHaveAttribute('aria-hidden', 'true')
  })

  test('the grid fills the whole band, not one of the thirds', async () => {
    const widths = await session.window.evaluate(() => {
      const band = document.querySelector('.band--header') as HTMLElement | null
      const grid = document.querySelector('.band--header .band__grid') as HTMLElement | null
      if (band === null || grid === null) return null
      return { band: band.getBoundingClientRect().width, grid: grid.getBoundingClientRect().width }
    })

    expect(widths).not.toBeNull()
    expect(widths!.grid).toBeGreaterThan(widths!.band * 0.95)
  })
})
