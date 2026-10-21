import { execFile } from 'node:child_process'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithLongTable } from './fixtures.js'
import { hasPdftotext } from './external-tools.js'

/**
 * The editor paginates live: the sheet count responds to typing, and the long table's last row
 * stays inside the paper. Font and the exact break line vary by machine, and are not pinned.
 */
test.describe('live pagination', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-pagination-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  test('a long table ends inside the last sheet', async () => {
    const source = join(directory, 'tabela-longa.docx')
    await writeFile(source, await docxWithLongTable())
    await stubDialogs(session.app, { open: source, messageBox: 1 })
    await menu(session, 'open')
    const sheets = session.window.locator('.paper')
    const rows = session.window.locator('.ProseMirror table').first().locator('tr')
    await expect(rows).toHaveCount(80)
    await expect.poll(() => sheets.count()).toBeGreaterThanOrEqual(2)
    await expect(async () => {
      const row = await rows.last().boundingBox()
      const sheet = await sheets.last().boundingBox()
      expect(row).not.toBeNull()
      expect(sheet).not.toBeNull()
      expect(row!.x).toBeGreaterThanOrEqual(sheet!.x)
      expect(row!.y).toBeGreaterThanOrEqual(sheet!.y)
      expect(row!.x + row!.width).toBeLessThanOrEqual(sheet!.x + sheet!.width)
      expect(row!.y + row!.height).toBeLessThanOrEqual(sheet!.y + sheet!.height)
    }).toPass()
  })

  test("the imported table cell uses Word's margin, and the table fits the paper sheets", async () => {
    // Without `w:tblCellMar` Word uses 0 top and bottom: the eighty rows take two sheets, as in
    // LibreOffice.
    const source = join(directory, 'tabela-longa.docx')
    await writeFile(source, await docxWithLongTable())
    await stubDialogs(session.app, { open: source, messageBox: 1 })
    await menu(session, 'open')
    const célula = session.window.locator('.ProseMirror td').first()
    await expect(célula).toBeVisible()
    expect(await célula.evaluate((cell) => getComputedStyle(cell).paddingTop)).toBe('0px')
    expect(await célula.evaluate((cell) => getComputedStyle(cell).paddingLeft)).toBe('7.2px')
    await expect(session.window.locator('.paper')).toHaveCount(2)
  })

  test('a new document has a single sheet', async () => {
    await menu(session, 'new-document')
    await expect(session.window.locator('.paper')).toHaveCount(1)
    await expect(session.window.locator('.statusbar__metric', { hasText: 'página' })).toContainText(
      '1 página',
    )
  })

  test('a manual page break opens a new sheet', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Capa.')

    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda folha.')

    await expect(session.window.locator('.paper')).toHaveCount(2)
  })

  test('text that does not fit is pushed to the next sheet', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()

    // By keyboard, not in the model: typing is what creates the sheet.
    for (let i = 0; i < 60; i++) {
      await session.window.keyboard.type(`Linha ${i} de um parágrafo qualquer para ocupar a folha.`)
      await session.window.keyboard.press('Enter')
    }

    await expect(session.window.locator('.paper')).not.toHaveCount(1)
  })

  test('each sheet repeats the header with its own number', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Primeira.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda.')

    // The number beside the sheet is what says "this is page 2".
    await expect(session.window.locator('.paper__number')).toHaveText(['1', '2'])
  })

  test('the paper comes out with the same sheets the screen shows', async () => {
    // The PDF page count is the screen's.
    const destinationFile = join(directory, 'saida.pdf')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })

    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Primeira folha.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda folha.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Terceira folha.')

    // Wait for the sheets instead of counting them: pagination settles one frame after the last
    // key, and `count()` does not ask again.
    const paperSheets = session.window.locator('.paper')
    await expect(paperSheets).toHaveCount(3)
    const onScreen = await paperSheets.count()

    await menu(session, 'export-pdf')
    await expect.poll(async () => countPages(destinationFile), { timeout: 30000 }).toBe(onScreen)
  })

  test('a paragraph that does not fit is cut between lines, at the same place on screen and in the PDF', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext not installed')
    const destinationFile = join(directory, 'corte.pdf')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await paragraphAcrossSheet(session)

    const paperSheets = session.window.locator('.paper')
    await expect(paperSheets).toHaveCount(2)
    await expect.poll(() => screenCut(session)).not.toBeNull()
    const cut = (await screenCut(session))!

    // On screen: the line before the spacer ends inside the first sheet, and the one after starts
    // inside the second.
    const limits = await session.window.evaluate((token) => {
      const paragraph = document.querySelector('.ProseMirror .page-line-gap')!.closest('p')!
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
      for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
        const at = ` ${node.textContent ?? ''} `.indexOf(` ${token} `)
        if (at === -1) continue
        const range = document.createRange()
        range.setStart(node, at)
        range.setEnd(node, at + token.length)
        return range.getBoundingClientRect().top
      }
      return null
    }, cut.afterCut)
    const sheets = await paperSheets.evaluateAll((all) =>
      all.map((sheet) => sheet.getBoundingClientRect().top),
    )
    expect(limits).not.toBeNull()
    expect(limits!).toBeGreaterThan(sheets[1]!)

    await menu(session, 'export-pdf')
    await expect.poll(async () => countPages(destinationFile), { timeout: 30000 }).toBe(2)
    const first = await wordsOfPage(destinationFile, 1)
    const second = await wordsOfPage(destinationFile, 2)
    expect(first.at(-1)).toBe(cut.beforeCut)
    expect(second[0]).toBe(cut.afterCut)
  })

  test('widows and orphans: an orphan at the foot moves the whole paragraph to the next sheet', async () => {
    // The filler leaves a single line at the foot: with widow control, the paragraph moves down
    // whole.
    await paragraphAcrossSheet(session, 30, 100)
    await expect(session.window.locator('.paper')).toHaveCount(2)
    await expect.poll(() => linesAroundCut(session)).toEqual({ beforeCut: 0, afterCut: 0 })
  })

  test('widows and orphans: the cut leaves at least two lines on each side', async () => {
    // Twenty-eight: the break takes the second-to-last line along with the last. A session of its
    // own, because switching documents with the first one dirty opens the discard prompt.
    await paragraphAcrossSheet(session, 28, 100)
    // In `toPass`: under load the measurement is still settling.
    await expect(async () => {
      const surroundingLines = await linesAroundCut(session)
      expect(surroundingLines.beforeCut).toBeGreaterThanOrEqual(2)
      expect(surroundingLines.afterCut).toBeGreaterThanOrEqual(2)
    }).toPass()
  })

  test('the table header row repeats at the top of each sheet, on screen and in the PDF', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext not installed')
    const source = join(directory, 'cabecalho.docx')
    const destinationFile = join(directory, 'cabecalho.pdf')
    await writeFile(source, await docxWithLongTable(80, true))
    await stubDialogs(session.app, { open: source, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')

    const sheets = session.window.locator('.paper')
    await expect.poll(() => sheets.count()).toBeGreaterThanOrEqual(2)
    const repeated = session.window.locator('.page-repeated-header')
    await expect.poll(() => repeated.count()).toBe((await sheets.count()) - 1)

    // The copy stays inside the second sheet, above its first row.
    const copy = await repeated.first().boundingBox()
    const paperSheet = await sheets.nth(1).boundingBox()
    expect(copy!.y).toBeGreaterThanOrEqual(paperSheet!.y)
    expect(copy!.y + copy!.height).toBeLessThanOrEqual(paperSheet!.y + paperSheet!.height)
    await expect(repeated.first()).toContainText('Cabeçalho repetido')
    // And aligned with the original table, column by column.
    const original = await session.window.locator('.page__content th').first().boundingBox()
    const clone = await repeated.first().locator('th').first().boundingBox()
    expect(Math.abs(clone!.x - original!.x)).toBeLessThanOrEqual(1)
    expect(Math.abs(clone!.width - original!.width)).toBeLessThanOrEqual(1)
    // Inside the gap of the row opening the sheet, not over its text.
    const neighbor = await session.window.evaluate(() => {
      const header = document.querySelector('.page-repeated-header')!
      const row = header.closest('tr')!
      return {
        copy: header.getBoundingClientRect().bottom,
        line: row.cells[0]!.getBoundingClientRect().top,
      }
    })
    expect(neighbor.copy).toBeGreaterThan(neighbor.line)

    await menu(session, 'export-pdf')
    await expect.poll(async () => countPages(destinationFile), { timeout: 30000 }).toBe(await sheets.count())
    const second = await wordsOfPage(destinationFile, 2)
    expect(second.slice(0, 2)).toEqual(['Cabeçalho', 'repetido'])

    // Turning the header row off from the table menu removes the repetition.
    await session.window
      .locator('.page__content th', { hasText: 'Cabeçalho repetido' })
      .first()
      .click({ button: 'right' })
    await session.window.getByRole('menuitem', { name: 'Linha de cabeçalho' }).click()
    await expect(repeated).toHaveCount(0)
  })

  test('zoom enlarges the sheet without changing where the page cuts', async () => {
    await paragraphAcrossSheet(session)
    const paperSheets = session.window.locator('.paper')
    await expect(paperSheets).toHaveCount(2)
    await expect.poll(() => screenCut(session)).not.toBeNull()
    const cut = await screenCut(session)
    const width = (await paperSheets.first().boundingBox())!.width

    const level = session.window.locator('.statusbar__zoom-level')
    await expect(level).toHaveText('100%')
    await session.window.getByRole('button', { name: 'Ampliar' }).click()
    await menu(session, 'zoom-in')
    await expect(level).toHaveText('125%')

    // The sheet grows on screen; pagination, measured at 100 %, stays where it was.
    await expect
      .poll(async () => (await paperSheets.first().boundingBox())!.width)
      .toBeCloseTo(width * 1.25, 0)
    await expect(paperSheets).toHaveCount(2)
    expect(await screenCut(session)).toEqual(cut)

    await menu(session, 'zoom-out')
    await expect(level).toHaveText('110%')
    await session.window.getByRole('button', { name: 'Ajustar à largura' }).click()
    await expect(session.window.getByRole('button', { name: 'Ajustar à largura' })).toHaveAttribute(
      'aria-pressed',
      'true',
    )
    expect(await screenCut(session)).toEqual(cut)

    await menu(session, 'zoom-reset')
    await expect(level).toHaveText('100%')
    await expect.poll(async () => (await paperSheets.first().boundingBox())!.width).toBeCloseTo(width, 0)
  })

  test('the cut between lines comes back when the text above it shrinks', async () => {
    // An already drawn spacer does not enter the next measurement, or the break would not come
    // back.
    await paragraphAcrossSheet(session)
    await expect.poll(() => screenCut(session)).not.toBeNull()
    const original = await screenCut(session)
    // A separate history step: together, undo would take the whole paragraph.
    await session.window.waitForTimeout(700)

    const paragraphStart = async (): Promise<void> => {
      await session.window.evaluate(() => {
        const paragraph = Array.from(document.querySelectorAll('.ProseMirror > p')).at(-1)!
        const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
        const text = walker.nextNode()!
        window.getSelection()!.collapse(text, 0)
      })
      // ProseMirror reads the DOM selection on `selectionchange`, a frame later.
      await session.window.waitForTimeout(100)
    }

    await paragraphStart()
    await session.window.keyboard.insertText(`${Array.from({ length: 25 }, () => 'extra').join(' ')} `)
    await expect.poll(() => screenCut(session)).not.toEqual(original)

    await session.window.keyboard.press('Control+z')
    await expect.poll(() => screenCut(session)).toEqual(original)

    await paragraphStart()
    for (let i = 0; i < 3; i++) await session.window.keyboard.press('Enter')
    for (let i = 0; i < 3; i++) await session.window.keyboard.press('Backspace')
    await expect.poll(() => screenCut(session)).toEqual(original)
  })
})

/**
 * Fills the sheet with short paragraphs and writes a long one of numbered words (`p1 p2 …`) that
 * crosses the sheet foot: unique words compare the screen break with the paper one without
 * depending on the font.
 */
async function paragraphAcrossSheet(session: Session, padding = 28, total = 160): Promise<void> {
  await menu(session, 'new-document')
  await session.window.locator('.ProseMirror').click()
  for (let i = 0; i < padding; i++) {
    await session.window.keyboard.insertText(`Enchimento ${i}.`)
    await session.window.keyboard.press('Enter')
  }
  const tokens = Array.from({ length: total }, (_, index) => `p${index + 1}`).join(' ')
  await session.window.keyboard.insertText(tokens)
}

/**
 * How many lines of the split paragraph sit before and after the spacer; zero and zero without a
 * break.
 */
async function linesAroundCut(session: Session): Promise<{ beforeCut: number; afterCut: number }> {
  return session.window.evaluate(() => {
    const gap = document.querySelector('.ProseMirror .page-line-gap')
    const paragraph = gap?.closest('p')
    if (gap === null || gap === undefined || paragraph === null || paragraph === undefined) {
      return { beforeCut: 0, afterCut: 0 }
    }
    const lines = (range: Range) =>
      new Set(
        Array.from(range.getClientRects())
          .filter((rect) => rect.height > 0 && rect.width > 0)
          .map((rect) => Math.round(rect.top)),
      ).size
    const before = document.createRange()
    before.setStart(paragraph, 0)
    before.setEndBefore(gap)
    const after = document.createRange()
    after.setStartAfter(gap)
    after.setEnd(paragraph, paragraph.childNodes.length)
    return { beforeCut: lines(before), afterCut: lines(after) }
  })
}

/** The last word before the spacer and the first after it, read from the DOM. */
async function screenCut(session: Session): Promise<{ beforeCut: string; afterCut: string } | null> {
  return session.window.evaluate(() => {
    const gap = document.querySelector('.ProseMirror .page-line-gap')
    const paragraph = gap?.closest('p')
    if (gap === null || gap === undefined || paragraph === null || paragraph === undefined) return null
    const before = document.createRange()
    before.setStart(paragraph, 0)
    before.setEndBefore(gap)
    const after = document.createRange()
    after.setStartAfter(gap)
    after.setEnd(paragraph, paragraph.childNodes.length)
    const words = (range: Range) => range.toString().trim().split(/\s+/)
    return { beforeCut: words(before).at(-1) ?? '', afterCut: words(after)[0] ?? '' }
  })
}

/** The words on a PDF page, through the system `pdftotext`. */
async function wordsOfPage(filePath: string, page: number): Promise<string[]> {
  const { stdout } = await promisify(execFile)('pdftotext', [
    '-f',
    String(page),
    '-l',
    String(page),
    '-layout',
    filePath,
    '-',
  ])
  return stdout.split(/\s+/).filter((word) => word.length > 0)
}

/**
 * PDF pages by their `/Type /Page` objects, without a library; `[^s]` tells `/Page` from `/Pages`.
 */
async function countPages(filePath: string): Promise<number> {
  try {
    const bytes = await readFile(filePath)
    return (bytes.toString('latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length
  } catch {
    // File not written yet: `poll` tries again.
    return 0
  }
}
