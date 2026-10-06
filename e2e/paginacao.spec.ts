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
test.describe('paginação ao vivo', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-paginacao-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('tabela longa termina dentro da última folha', async () => {
    const source = join(pasta, 'tabela-longa.docx')
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

  test('a célula da tabela importada usa a margem do Word, e a tabela cabe nas folhas do papel', async () => {
    // Without `w:tblCellMar` Word uses 0 top and bottom: the eighty rows take two sheets, as in
    // LibreOffice.
    const source = join(pasta, 'tabela-longa.docx')
    await writeFile(source, await docxWithLongTable())
    await stubDialogs(session.app, { open: source, messageBox: 1 })
    await menu(session, 'open')
    const célula = session.window.locator('.ProseMirror td').first()
    await expect(célula).toBeVisible()
    expect(await célula.evaluate((cell) => getComputedStyle(cell).paddingTop)).toBe('0px')
    expect(await célula.evaluate((cell) => getComputedStyle(cell).paddingLeft)).toBe('7.2px')
    await expect(session.window.locator('.paper')).toHaveCount(2)
  })

  test('documento novo tem uma folha só', async () => {
    await menu(session, 'new-document')
    await expect(session.window.locator('.paper')).toHaveCount(1)
    await expect(session.window.locator('.statusbar__metric', { hasText: 'página' })).toContainText(
      '1 página',
    )
  })

  test('a quebra de página pedida à mão abre folha nova', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Capa.')

    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda folha.')

    await expect(session.window.locator('.paper')).toHaveCount(2)
  })

  test('texto que não cabe empurra para a folha seguinte', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()

    // By keyboard, not in the model: typing is what creates the sheet.
    for (let i = 0; i < 60; i++) {
      await session.window.keyboard.type(`Linha ${i} de um parágrafo qualquer para ocupar a folha.`)
      await session.window.keyboard.press('Enter')
    }

    await expect(session.window.locator('.paper')).not.toHaveCount(1)
  })

  test('cada folha repete o cabeçalho com o número dela', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Primeira.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda.')

    // The number beside the sheet is what says "this is page 2".
    await expect(session.window.locator('.paper__number')).toHaveText(['1', '2'])
  })

  test('o papel sai com as mesmas folhas que a tela mostra', async () => {
    // The PDF page count is the screen's.
    const destino = join(pasta, 'saida.pdf')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })

    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Primeira folha.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda folha.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Terceira folha.')

    // Wait for the sheets instead of counting them: pagination settles one frame after the last
    // key, and `count()` does not ask again.
    const folhas = session.window.locator('.paper')
    await expect(folhas).toHaveCount(3)
    const naTela = await folhas.count()

    await menu(session, 'export-pdf')
    await expect.poll(async () => contarPaginas(destino), { timeout: 30000 }).toBe(naTela)
  })

  test('o parágrafo que não cabe é cortado entre linhas, no mesmo lugar na tela e no PDF', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext não instalado')
    const destino = join(pasta, 'corte.pdf')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await paragrafoAtravessandoAFolha(session)

    const folhas = session.window.locator('.paper')
    await expect(folhas).toHaveCount(2)
    await expect.poll(() => corteNaTela(session)).not.toBeNull()
    const corte = (await corteNaTela(session))!

    // On screen: the line before the spacer ends inside the first sheet, and the one after starts
    // inside the second.
    const limites = await session.window.evaluate((palavra) => {
      const paragraph = document.querySelector('.ProseMirror .page-line-gap')!.closest('p')!
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
      for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
        const at = ` ${node.textContent ?? ''} `.indexOf(` ${palavra} `)
        if (at === -1) continue
        const range = document.createRange()
        range.setStart(node, at)
        range.setEnd(node, at + palavra.length)
        return range.getBoundingClientRect().top
      }
      return null
    }, corte.depois)
    const sheets = await folhas.evaluateAll((all) => all.map((sheet) => sheet.getBoundingClientRect().top))
    expect(limites).not.toBeNull()
    expect(limites!).toBeGreaterThan(sheets[1]!)

    await menu(session, 'export-pdf')
    await expect.poll(async () => contarPaginas(destino), { timeout: 30000 }).toBe(2)
    const primeira = await palavrasDaPagina(destino, 1)
    const segunda = await palavrasDaPagina(destino, 2)
    expect(primeira.at(-1)).toBe(corte.antes)
    expect(segunda[0]).toBe(corte.depois)
  })

  test('viúvas e órfãs: a órfã no pé leva o parágrafo inteiro para a folha seguinte', async () => {
    // The filler leaves a single line at the foot: with widow control, the paragraph moves down
    // whole.
    await paragrafoAtravessandoAFolha(session, 30, 100)
    await expect(session.window.locator('.paper')).toHaveCount(2)
    await expect.poll(() => linhasEmVoltaDoCorte(session)).toEqual({ antes: 0, depois: 0 })
  })

  test('viúvas e órfãs: o corte deixa ao menos duas linhas de cada lado', async () => {
    // Twenty-eight: the break takes the second-to-last line along with the last. A session of its
    // own, because switching documents with the first one dirty opens the discard prompt.
    await paragrafoAtravessandoAFolha(session, 28, 100)
    // In `toPass`: under load the measurement is still settling.
    await expect(async () => {
      const linhas = await linhasEmVoltaDoCorte(session)
      expect(linhas.antes).toBeGreaterThanOrEqual(2)
      expect(linhas.depois).toBeGreaterThanOrEqual(2)
    }).toPass()
  })

  test('a linha de cabeçalho da tabela se repete no alto de cada folha, na tela e no PDF', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext não instalado')
    const source = join(pasta, 'cabecalho.docx')
    const destino = join(pasta, 'cabecalho.pdf')
    await writeFile(source, await docxWithLongTable(80, true))
    await stubDialogs(session.app, { open: source, save: destino, messageBox: 1 })
    await menu(session, 'open')

    const sheets = session.window.locator('.paper')
    await expect.poll(() => sheets.count()).toBeGreaterThanOrEqual(2)
    const repetidos = session.window.locator('.page-repeated-header')
    await expect.poll(() => repetidos.count()).toBe((await sheets.count()) - 1)

    // The copy stays inside the second sheet, above its first row.
    const copia = await repetidos.first().boundingBox()
    const folha = await sheets.nth(1).boundingBox()
    expect(copia!.y).toBeGreaterThanOrEqual(folha!.y)
    expect(copia!.y + copia!.height).toBeLessThanOrEqual(folha!.y + folha!.height)
    await expect(repetidos.first()).toContainText('Cabeçalho repetido')
    // And aligned with the original table, column by column.
    const original = await session.window.locator('.page__content th').first().boundingBox()
    const clone = await repetidos.first().locator('th').first().boundingBox()
    expect(Math.abs(clone!.x - original!.x)).toBeLessThanOrEqual(1)
    expect(Math.abs(clone!.width - original!.width)).toBeLessThanOrEqual(1)
    // Inside the gap of the row opening the sheet, not over its text.
    const vizinha = await session.window.evaluate(() => {
      const header = document.querySelector('.page-repeated-header')!
      const row = header.closest('tr')!
      return {
        copia: header.getBoundingClientRect().bottom,
        linha: row.cells[0]!.getBoundingClientRect().top,
      }
    })
    expect(vizinha.copia).toBeGreaterThan(vizinha.linha)

    await menu(session, 'export-pdf')
    await expect.poll(async () => contarPaginas(destino), { timeout: 30000 }).toBe(await sheets.count())
    const segunda = await palavrasDaPagina(destino, 2)
    expect(segunda.slice(0, 2)).toEqual(['Cabeçalho', 'repetido'])

    // Turning the header row off from the table menu removes the repetition.
    await session.window
      .locator('.page__content th', { hasText: 'Cabeçalho repetido' })
      .first()
      .click({ button: 'right' })
    await session.window.getByRole('menuitem', { name: 'Linha de cabeçalho' }).click()
    await expect(repetidos).toHaveCount(0)
  })

  test('o zoom aumenta a folha sem mudar onde a página corta', async () => {
    await paragrafoAtravessandoAFolha(session)
    const folhas = session.window.locator('.paper')
    await expect(folhas).toHaveCount(2)
    await expect.poll(() => corteNaTela(session)).not.toBeNull()
    const corte = await corteNaTela(session)
    const largura = (await folhas.first().boundingBox())!.width

    const nivel = session.window.locator('.statusbar__zoom-level')
    await expect(nivel).toHaveText('100%')
    await session.window.getByRole('button', { name: 'Ampliar' }).click()
    await menu(session, 'zoom-in')
    await expect(nivel).toHaveText('125%')

    // The sheet grows on screen; pagination, measured at 100 %, stays where it was.
    await expect.poll(async () => (await folhas.first().boundingBox())!.width).toBeCloseTo(largura * 1.25, 0)
    await expect(folhas).toHaveCount(2)
    expect(await corteNaTela(session)).toEqual(corte)

    await menu(session, 'zoom-out')
    await expect(nivel).toHaveText('110%')
    await session.window.getByRole('button', { name: 'Ajustar à largura' }).click()
    await expect(session.window.getByRole('button', { name: 'Ajustar à largura' })).toHaveAttribute(
      'aria-pressed',
      'true',
    )
    expect(await corteNaTela(session)).toEqual(corte)

    await menu(session, 'zoom-reset')
    await expect(nivel).toHaveText('100%')
    await expect.poll(async () => (await folhas.first().boundingBox())!.width).toBeCloseTo(largura, 0)
  })

  test('o corte entre linhas volta quando o texto acima dele encolhe', async () => {
    // An already drawn spacer does not enter the next measurement, or the break would not come
    // back.
    await paragrafoAtravessandoAFolha(session)
    await expect.poll(() => corteNaTela(session)).not.toBeNull()
    const original = await corteNaTela(session)
    // A separate history step: together, undo would take the whole paragraph.
    await session.window.waitForTimeout(700)

    const inicioDoParagrafo = async (): Promise<void> => {
      await session.window.evaluate(() => {
        const paragraph = Array.from(document.querySelectorAll('.ProseMirror > p')).at(-1)!
        const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
        const text = walker.nextNode()!
        window.getSelection()!.collapse(text, 0)
      })
      // ProseMirror reads the DOM selection on `selectionchange`, a frame later.
      await session.window.waitForTimeout(100)
    }

    await inicioDoParagrafo()
    await session.window.keyboard.insertText(`${Array.from({ length: 25 }, () => 'extra').join(' ')} `)
    await expect.poll(() => corteNaTela(session)).not.toEqual(original)

    await session.window.keyboard.press('Control+z')
    await expect.poll(() => corteNaTela(session)).toEqual(original)

    await inicioDoParagrafo()
    for (let i = 0; i < 3; i++) await session.window.keyboard.press('Enter')
    for (let i = 0; i < 3; i++) await session.window.keyboard.press('Backspace')
    await expect.poll(() => corteNaTela(session)).toEqual(original)
  })
})

/**
 * Fills the sheet with short paragraphs and writes a long one of numbered words (`p1 p2 …`) that
 * crosses the sheet foot: unique words compare the screen break with the paper one without
 * depending on the font.
 */
async function paragrafoAtravessandoAFolha(session: Session, enchimento = 28, total = 160): Promise<void> {
  await menu(session, 'new-document')
  await session.window.locator('.ProseMirror').click()
  for (let i = 0; i < enchimento; i++) {
    await session.window.keyboard.insertText(`Enchimento ${i}.`)
    await session.window.keyboard.press('Enter')
  }
  const palavras = Array.from({ length: total }, (_, index) => `p${index + 1}`).join(' ')
  await session.window.keyboard.insertText(palavras)
}

/**
 * How many lines of the split paragraph sit before and after the spacer; zero and zero without a
 * break.
 */
async function linhasEmVoltaDoCorte(session: Session): Promise<{ antes: number; depois: number }> {
  return session.window.evaluate(() => {
    const gap = document.querySelector('.ProseMirror .page-line-gap')
    const paragraph = gap?.closest('p')
    if (gap === null || gap === undefined || paragraph === null || paragraph === undefined) {
      return { antes: 0, depois: 0 }
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
    return { antes: lines(before), depois: lines(after) }
  })
}

/** The last word before the spacer and the first after it, read from the DOM. */
async function corteNaTela(session: Session): Promise<{ antes: string; depois: string } | null> {
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
    return { antes: words(before).at(-1) ?? '', depois: words(after)[0] ?? '' }
  })
}

/** The words on a PDF page, through the system `pdftotext`. */
async function palavrasDaPagina(caminho: string, pagina: number): Promise<string[]> {
  const { stdout } = await promisify(execFile)('pdftotext', [
    '-f',
    String(pagina),
    '-l',
    String(pagina),
    '-layout',
    caminho,
    '-',
  ])
  return stdout.split(/\s+/).filter((word) => word.length > 0)
}

/**
 * PDF pages by their `/Type /Page` objects, without a library; `[^s]` tells `/Page` from `/Pages`.
 */
async function contarPaginas(caminho: string): Promise<number> {
  try {
    const bytes = await readFile(caminho)
    return (bytes.toString('latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length
  } catch {
    // File not written yet: `poll` tries again.
    return 0
  }
}
