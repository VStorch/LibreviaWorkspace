import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithPageNumbering, entryOf } from './fixtures.js'
import { hasPdftotext, pdfText } from './external-tools.js'

/**
 * Page numbering: each sheet's number with the start and format the document asks for, the field
 * the user inserted, and the band switches.
 */
test.describe('page numbering', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-numbering-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  const footers = (session: Session) =>
    session.window
      .locator('.band--footer')
      .evaluateAll((bands) => bands.map((band) => (band.textContent ?? '').trim()))

  async function openDocument(): Promise<string> {
    const sourceFile = join(directory, 'entrada.docx')
    await writeFile(sourceFile, await docxWithPageNumbering())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.paper')).toHaveCount(2)
    return sourceFile
  }

  test('each sheet shows the number in the format and from the start of the document', async () => {
    await openDocument()
    await expect.poll(() => footers(session)).toEqual(['Página iii', 'Página iv'])
  })

  test('the PDF numbers the sheets like the screen', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext not installed')
    await openDocument()
    const destinationFile = join(directory, 'saida.pdf')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'export-pdf')
    await expect
      .poll(() => pdfText(destinationFile), { timeout: 30_000 })
      .toMatch(/Página iii[\s\S]*Página iv/)
  })

  test("a field inserted in Word's band becomes a field in the file and counts on screen", async () => {
    await openDocument()
    const piece = session.window.locator('.band--footer .band__text').first()
    await piece.click()
    await session.window.keyboard.press('Home')
    await session.window.getByRole('button', { name: 'Inserir total de páginas' }).click()
    await session.window.keyboard.type(' folhas — ')
    await session.window.locator('.ProseMirror').click()
    await expect.poll(() => footers(session)).toEqual(['2 folhas — Página iii', '2 folhas — Página iv'])

    const destinationFile = join(directory, 'saida.docx')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const footer = await entryOf(destinationFile, 'word/footer1.xml')
    expect(footer).toMatch(/w:instr=" NUMPAGES "/)
    expect(footer).not.toContain('{total}')
  })

  test('different first page and start changed in the panel apply on screen and in the file', async () => {
    await openDocument()
    await menu(session, 'page-setup')
    const panel = session.window.getByRole('dialog', { name: 'Configuração de página' })
    await panel.getByLabel('Primeira página diferente').check()
    await panel.getByLabel('Pares e ímpares diferentes').check()
    await panel.getByLabel('Formato do número').selectOption('upperRoman')
    await panel.getByLabel('Começar em').fill('1')
    await panel.getByRole('button', { name: 'Aplicar' }).click()

    // A title page without a footer; sheet 2 is page II, even, and without its own even footer.
    await expect.poll(() => footers(session)).toEqual([])

    const destinationFile = join(directory, 'saida.docx')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const document = await entryOf(destinationFile, 'word/document.xml')
    expect(document).toContain('<w:titlePg')
    expect(document).toMatch(/<w:pgNumType w:fmt="upperRoman" w:start="1"/)
    expect(await entryOf(destinationFile, 'word/settings.xml')).toContain('w:evenAndOddHeaders')
  })

  test('new document: a page number inserted in the footer shows on each sheet', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Primeira.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda.')
    await session.window.getByRole('button', { name: 'Inserir número da página' }).click()
    await expect.poll(() => footers(session)).toEqual(['1', '2'])
  })
})
