import { execFile } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithColumns, docxWithSections, entryOf } from './fixtures.js'
import { hasPdfinfo } from './external-tools.js'

/**
 * Sections: each sheet with its section's paper, band and number.
 *
 * The document has portrait (Roman), landscape (restarting at 1) and portrait again starting on an
 * odd page; Word inserts a blank sheet before it, already with the new section's paper. The last
 * two sections declare no footer, and inherit the first one's.
 */
test.describe('sections', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-sections-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  async function openDocument(): Promise<void> {
    const sourceFile = join(directory, 'secoes.docx')
    await writeFile(sourceFile, await docxWithSections())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.paper')).toHaveCount(4)
  }

  test('portrait and landscape in the same stack, with the blank sheet of the odd section', async () => {
    await openDocument()
    const sheets = await session.window
      .locator('.paper')
      .evaluateAll((papers) => papers.map((paper) => paper.getBoundingClientRect()))

    // The second sheet is wider than tall; the others, portrait.
    expect(sheets[0]!.width).toBeLessThan(sheets[0]!.height)
    expect(sheets[1]!.width).toBeGreaterThan(sheets[1]!.height)
    expect(sheets[2]!.width).toBeLessThan(sheets[2]!.height)
    // Centered in the stack: the portrait one starts further right than the landscape one.
    expect(sheets[0]!.left).toBeGreaterThan(sheets[1]!.left)

    // The landscape section's text takes its width, not the portrait one's.
    const landscape = session.window.locator('.ProseMirror p', { hasText: 'Folha em paisagem.' })
    const box = await landscape.boundingBox()
    expect(box!.x).toBeLessThan(sheets[0]!.left)
    expect(box!.y).toBeGreaterThan(sheets[1]!.top)
    expect(box!.y).toBeLessThan(sheets[1]!.bottom)

    // The third sheet is the blank one; portrait again opens the fourth.
    const portrait = await session.window
      .locator('.ProseMirror p', { hasText: 'Retrato outra vez' })
      .boundingBox()
    expect(portrait!.y).toBeGreaterThan(sheets[3]!.top)
  })

  test('the inherited footer numbers each section with its own format and restart', async () => {
    await openDocument()
    await expect
      .poll(() =>
        session.window
          .locator('.band--footer')
          .evaluateAll((bands) => bands.map((band) => (band.textContent ?? '').trim())),
      )
      .toEqual(['Página i', 'Página 1', 'Página 2', 'Página 3'])
  })

  test("the PDF comes out with each section's paper and the same numbers", async () => {
    test.skip(!(await hasPdfinfo()), 'pdfinfo not installed')
    await openDocument()
    const destinationFile = join(directory, 'secoes.pdf')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'export-pdf')

    await expect.poll(() => sizes(destinationFile), { timeout: 30_000 }).toHaveLength(4)
    const [first, second, third, fourth] = await sizes(destinationFile)
    expect(first).toEqual('retrato')
    expect(second).toEqual('paisagem')
    expect(third).toEqual('retrato')
    expect(fourth).toEqual('retrato')

    const { stdout } = await promisify(execFile)('pdftotext', ['-layout', destinationFile, '-'])
    expect(stdout).toMatch(/Página i[\s\S]*Folha em paisagem[\s\S]*Página 1[\s\S]*Página 2[\s\S]*Página 3/)
  })

  const orientations = (session: Session) =>
    session.window.locator('.paper').evaluateAll((papers) =>
      papers.map((paper) => {
        const box = paper.getBoundingClientRect()
        return box.width > box.height ? 'paisagem' : 'retrato'
      }),
    )

  test('insert a section break, turn only the lower section and delete the break', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Em retrato.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Em paisagem.')
    await session.window.keyboard.press('Home')
    await menu(session, 'insert-section-next-page')
    await expect(session.window.locator('.paper')).toHaveCount(2)

    // The cursor is in the lower section: "this section" means only that one.
    await session.window.locator('.ProseMirror p', { hasText: 'Em paisagem.' }).click()
    await menu(session, 'page-setup')
    const panel = session.window.getByRole('dialog', { name: 'Configuração de página' })
    await panel.getByLabel('Orientação').selectOption('landscape')
    await expect(panel.getByLabel('Nesta seção')).toBeChecked()
    await panel.getByRole('button', { name: 'Aplicar' }).click()
    await expect.poll(() => orientations(session)).toEqual(['retrato', 'paisagem'])

    const destinationFile = join(directory, 'duas-secoes.docx')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const body = await entryOf(destinationFile, 'word/document.xml')
    expect(body.match(/<w:sectPr/g)).toHaveLength(2)
    expect(body).toMatch(/<w:p>(?:(?!<\/w:p>).)*<w:sectPr(?:(?!<\/w:sectPr>).)*w:h="16838"/)
    expect(body).toMatch(/<w:sectPr(?:(?!<\/w:sectPr>).)*w:orient="landscape"(?:(?!<w:sectPr).)*<\/w:body>/)

    // With the break deleted, the range above takes the format of the one below, as in Word.
    await menu(session, 'delete-section-break')
    await expect.poll(() => orientations(session)).toEqual(['paisagem'])
  })

  test('unlinking the footer of the landscape section gives it a footer of its own', async () => {
    await openDocument()
    await session.window.locator('.ProseMirror p', { hasText: 'Folha em paisagem.' }).click()
    await menu(session, 'page-setup')
    const panel = session.window.getByRole('dialog', { name: 'Configuração de página' })
    const link = panel.getByLabel('Rodapé: vincular ao anterior')
    await expect(link).toBeChecked()
    await link.uncheck()
    await panel.getByRole('button', { name: 'Aplicar' }).click()

    // The copy is the landscape section's: editing it does not change the first one's footer.
    const piece = session.window.locator('.band--footer .band__text').nth(1)
    await piece.click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.type('Anexo — ')
    await session.window.locator('.ProseMirror').click()
    await expect
      .poll(() =>
        session.window
          .locator('.band--footer')
          .evaluateAll((bands) => bands.map((band) => (band.textContent ?? '').trim())),
      )
      .toEqual(['Página i', 'Anexo — Página 1', 'Anexo — Página 2', 'Anexo — Página 3'])

    const destinationFile = join(directory, 'desvinculado.docx')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(destinationFile, 'word/footer1.xml')).not.toContain('Anexo')
    const body = await entryOf(destinationFile, 'word/document.xml')
    expect(body.match(/<w:footerReference/g)).toHaveLength(2)
  })

  test('balanced columns before the continuous section, on screen and in the PDF', async () => {
    const sourceFile = join(directory, 'colunas.docx')
    await writeFile(sourceFile, await docxWithColumns())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.paper')).toHaveCount(1)

    const box = (text: string) => session.window.locator('.ProseMirror p', { hasText: text }).boundingBox()
    // Six paragraphs in two balanced columns: three on each side.
    await expect
      .poll(async () => (await box('Parágrafo 4'))!.x)
      .toBeGreaterThan((await box('Parágrafo 1'))!.x + 100)
    const first = (await box('Parágrafo 1'))!
    const fourth = (await box('Parágrafo 4'))!
    expect(Math.abs(fourth.y - first.y)).toBeLessThan(2)
    expect(fourth.width).toBeLessThan(first.width + 1)
    // The lower section's text moves down to the foot of the tallest column, and takes the whole
    // sheet.
    const third = (await box('Parágrafo 3'))!
    const after = (await box('Depois das colunas'))!
    expect(after.y).toBeGreaterThan(third.y + third.height - 1)
    expect(after.width).toBeGreaterThan(first.width * 1.5)
    await expect(session.window.locator('.paper-column-line')).toHaveCount(1)

    test.skip(!(await hasPdfinfo()), 'pdftotext not installed')
    const destinationFile = join(directory, 'colunas.pdf')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'export-pdf')
    await expect.poll(() => sizes(destinationFile), { timeout: 30_000 }).toHaveLength(1)
    const { stdout } = await promisify(execFile)('pdftotext', ['-layout', destinationFile, '-'])
    expect(stdout).toMatch(/Parágrafo 1 em colunas\.\s+Parágrafo 4 em colunas\./)
  })

  test('Format → Columns and the column break in a new document', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Esquerda.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Direita.')
    await menu(session, 'format-columns')
    const dialog = session.window.getByRole('dialog', { name: 'Colunas' })
    await dialog.getByLabel('Número de colunas').fill('2')
    await dialog.getByRole('button', { name: 'Aplicar' }).click()

    await session.window.locator('.ProseMirror p', { hasText: 'Esquerda.' }).click()
    await session.window.keyboard.press('End')
    await menu(session, 'insert-column-break')
    const left = session.window.locator('.ProseMirror p', { hasText: 'Esquerda.' })
    const right = session.window.locator('.ProseMirror p', { hasText: 'Direita.' })
    await expect
      .poll(async () => (await right.boundingBox())!.x - (await left.boundingBox())!.x)
      .toBeGreaterThan(100)

    const destinationFile = join(directory, 'colunas-novo.docx')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const body = await entryOf(destinationFile, 'word/document.xml')
    expect(body).toMatch(/<w:cols [^>]*w:num="2"/)
    expect(body).toContain('w:type="column"')
  })

  test('undoing the section break undoes the section, on screen and in the file', async () => {
    // The section lives in the store library, and the text says whether it counts: undo removes the
    // mark and, with it, the section and the start the lower one gained.
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Antes.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Depois.')
    await session.window.keyboard.press('Home')
    await menu(session, 'insert-section-odd-page')
    await expect(session.window.locator('.paper')).toHaveCount(3)

    await session.window.keyboard.press('Control+z')
    await expect(session.window.locator('.paper')).toHaveCount(1)

    const destinationFile = join(directory, 'desfeito.docx')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const body = await entryOf(destinationFile, 'word/document.xml')
    expect(body.match(/<w:sectPr/g)).toHaveLength(1)
    expect(body).not.toContain('oddPage')
  })

  test('a draft from before sections refuses break and columns, saying why', async () => {
    // Saved along the old path, it would not take the break to the file: the change would show on
    // screen and vanish from the .docx.
    const draft = join(directory, 'antigo.sdoc')
    await writeFile(
      draft,
      JSON.stringify({
        format: 'sdoc',
        version: 5,
        page: { size: 'A4', orientation: 'portrait', margins: { top: 25, right: 25, bottom: 25, left: 25 } },
        doc: {
          type: 'doc',
          content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Rascunho antigo.' }] }],
        },
      }),
    )
    await stubDialogs(session.app, { open: draft, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Rascunho antigo.')
    await session.window.locator('.ProseMirror').click()

    await menu(session, 'insert-section-next-page')
    await expect(session.window.getByRole('alert')).toContainText('versão anterior')
    await expect(session.window.locator('.paper')).toHaveCount(1)
    await expect(session.window.locator('.ProseMirror p')).toHaveCount(1)
  })

  test('a section break in a list item stays in the item, and the list continues on the next sheet', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('- Primeiro item.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Segundo item.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Terceiro item.')
    await session.window.locator('.ProseMirror li', { hasText: 'Segundo item.' }).click()
    await menu(session, 'insert-section-next-page')
    await expect(session.window.locator('.paper')).toHaveCount(2)

    const sheets = await session.window
      .locator('.paper')
      .evaluateAll((papers) => papers.map((paper) => paper.getBoundingClientRect().top))
    const second = await session.window.locator('.ProseMirror li', { hasText: 'Segundo item.' }).boundingBox()
    const third = await session.window.locator('.ProseMirror li', { hasText: 'Terceiro item.' }).boundingBox()
    expect(second!.y).toBeLessThan(sheets[1]!)
    expect(third!.y).toBeGreaterThan(sheets[1]!)
  })
})

/** Each PDF sheet's orientation, by the size `pdfinfo` reports. */
async function sizes(filePath: string): Promise<string[]> {
  try {
    const { stdout } = await promisify(execFile)('pdfinfo', ['-f', '1', '-l', '99', filePath])
    return [...stdout.matchAll(/Page\s+\d+ size:\s+([\d.]+) x ([\d.]+)/g)].map((measure) =>
      Number(measure[1]) > Number(measure[2]) ? 'paisagem' : 'retrato',
    )
  } catch {
    return []
  }
}
