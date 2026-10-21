import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithMultilevelList, entryOf } from './fixtures.js'
import { hasPdftotext, pdfText } from './external-tools.js'

/**
 * Multilevel lists with the label Word would draw: the second level composes the first (`1.a)`),
 * the list continues after a paragraph (`3.`) and restarts (`10.`), which a CSS counter does not
 * do.
 */
test.describe('multilevel lists', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-lists-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  const marks = (session: Session) =>
    session.window
      .locator('.ProseMirror li[data-label]')
      .evaluateAll((items) => items.map((item) => item.getAttribute('data-label')))

  async function openMultilevel(): Promise<string> {
    const sourceFile = join(directory, 'entrada.docx')
    await writeFile(sourceFile, await docxWithMultilevelList())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Dez')
    return sourceFile
  }

  test('the screen numbers like Word: compound levels, continuation and restart', async () => {
    await openMultilevel()
    await expect.poll(() => marks(session)).toEqual(['1.', '1.a)', '1.b)', '2.', '3.', '10.'])

    // And it is the visible label: the paragraph's `::before` draws the item's.
    const drawn = await session.window
      .locator('.ProseMirror li[data-label] > p')
      .nth(1)
      .evaluate((paragraph) => getComputedStyle(paragraph, '::before').content)
    expect(drawn).toBe('"1.a)"')
  })

  test('the PDF shows the same marks as the screen', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext not installed')
    await openMultilevel()
    const destinationFile = join(directory, 'saida.pdf')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'export-pdf')

    await expect.poll(() => pdfText(destinationFile), { timeout: 30_000 }).toContain('1.a)')
    const text = await pdfText(destinationFile)
    expect(text).toMatch(/1\.b\)\s+Um-b/)
    expect(text).toMatch(/3\.\s+Três/)
    expect(text).toMatch(/10\.\s+Dez/)
  })

  test('opening and saving without editing returns numbering.xml byte for byte', async () => {
    const sourceFile = await openMultilevel()
    const destinationFile = join(directory, 'saida.docx')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    expect(await entryOf(destinationFile, 'word/numbering.xml')).toBe(
      await entryOf(sourceFile, 'word/numbering.xml'),
    )
  })

  /**
   * Saves, and opens a **copy** of what was saved: reopening the file itself would not guarantee
   * the screen came from disk rather than from the document already there.
   */
  async function saveAndReopen(): Promise<string> {
    const destinationFile = join(directory, `saida-${Date.now()}.docx`)
    const copy = join(directory, `copia-${Date.now()}.docx`)
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await copyFile(destinationFile, copy)
    await stubDialogs(session.app, { open: copy, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window).toHaveTitle(/copia/)
    await expect(session.window.locator('.ProseMirror')).toContainText('Dez')
    return destinationFile
  }

  async function rightClick(text: string, action: string): Promise<void> {
    await session.window
      .locator('.ProseMirror li p', { hasText: new RegExp(`^${text}$`) })
      .click({ button: 'right' })
    await session.window.getByRole('menuitem', { name: action }).click()
  }

  test('Restart at 1 from the context menu applies on screen and in the file', async () => {
    await openMultilevel()
    await rightClick('Três', 'Reiniciar em 1')
    await expect.poll(() => marks(session)).toEqual(['1.', '1.a)', '1.b)', '2.', '1.', '10.'])

    const destinationFile = await saveAndReopen()
    expect(await entryOf(destinationFile, 'word/numbering.xml')).toMatch(/<w:startOverride w:val="1" ?\/>/)
    await expect.poll(() => marks(session)).toEqual(['1.', '1.a)', '1.b)', '2.', '1.', '10.'])
  })

  test('Restart in the middle of the list splits the list there, as in Word', async () => {
    await openMultilevel()
    await rightClick('Um-b', 'Reiniciar em 1')
    await expect.poll(() => marks(session)).toEqual(['1.', '1.a)', '1.a)', '2.', '3.', '10.'])
  })

  test('Continue numbering joins the restarted list to the previous one', async () => {
    await openMultilevel()
    await rightClick('Dez', 'Continuar numeração')
    await expect.poll(() => marks(session)).toEqual(['1.', '1.a)', '1.b)', '2.', '3.', '4.'])

    await saveAndReopen()
    await expect.poll(() => marks(session)).toEqual(['1.', '1.a)', '1.b)', '2.', '3.', '4.'])
  })

  test('Set starting value starts the list at the requested number', async () => {
    await openMultilevel()
    await rightClick('Um', 'Definir valor inicial…')
    const field = session.window
      .getByRole('dialog', { name: 'Definir valor inicial' })
      .getByRole('spinbutton')
    await field.fill('5')
    await session.window
      .getByRole('dialog', { name: 'Definir valor inicial' })
      .getByRole('button', { name: 'Definir' })
      .click()
    // The rest of numbering 5 counts on its own, as in Word: the list above became another `w:num`,
    // with a restart.
    await expect.poll(() => marks(session)).toEqual(['5.', '5.a)', '5.b)', '6.', '1.', '10.'])

    await saveAndReopen()
    await expect.poll(() => marks(session)).toEqual(['5.', '5.a)', '5.b)', '6.', '1.', '10.'])
  })

  test('the gallery applies a ready multilevel list, and the file carries it', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('1. Capítulo')

    await session.window.getByRole('button', { name: 'Formato de lista…' }).click()
    const dialog = session.window.getByRole('dialog', { name: 'Formato de lista' })
    await dialog.getByRole('button', { name: 'I. A. 1. a.' }).click()
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect.poll(() => marks(session)).toEqual(['I.'])

    await session.window.locator('.ProseMirror').press('End')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.press('Tab')
    await session.window.keyboard.type('Seção')
    await expect.poll(() => marks(session)).toEqual(['I.', 'A.'])

    const destinationFile = join(directory, 'galeria.docx')
    await stubDialogs(session.app, { save: destinationFile, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const numbering = await entryOf(destinationFile, 'word/numbering.xml')
    expect(numbering).toContain('w:numFmt w:val="upperRoman"')
    expect(numbering).toContain('w:numFmt w:val="upperLetter"')
  })

  test('Tab demotes the item one level and Shift+Tab brings it back', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('1. Primeiro')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Segundo')
    await expect.poll(() => marks(session)).toEqual(['1.', '2.'])

    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('Tab')
    await expect.poll(() => marks(session)).toEqual(['1.', 'a.'])

    await session.window.keyboard.press('Shift+Tab')
    await expect.poll(() => marks(session)).toEqual(['1.', '2.'])
  })
})
