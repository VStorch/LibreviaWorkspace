import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithVerticalAlignment, entryOf } from './fixtures.js'

/**
 * Character and paragraph formatting across the whole seam: the button, the editor command and what
 * the sidecar writes. Superscript changes the meaning ("cm3" is not "cm³"): if it comes back from
 * the file, the path holds.
 */
test.describe('document formatting', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-formatting-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('the document superscript shows on screen and returns to the file', async () => {
    const sourceFile = join(folder, 'formula.docx')
    const destinationFile = join(folder, 'salva.docx')
    await writeFile(sourceFile, await docxWithVerticalAlignment())

    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('O ocupa')

    // On screen, `<sup>` and `<sub>`, which the PDF also carries.
    await expect(editor.locator('sup')).toHaveText('3')
    await expect(editor.locator('sub')).toHaveText('2')

    // Editing forces the writer to rewrite the paragraph, where the loss would happen.
    await editor.click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Mexido.')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const body = await entryOf(destinationFile, 'word/document.xml')
    expect(body).toContain('w:vertAlign w:val="superscript"')
    expect(body).toContain('w:vertAlign w:val="subscript"')

    // Nothing in the notice: the writer knows how to write this formatting.
    await expect(session.window.locator('.banner--notice')).toHaveCount(0)
  })

  test('the superscript, uppercase and small caps buttons mark the text', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('marcado')
    await session.window.keyboard.press('Control+a')

    for (const [label, selector] of [
      ['Sobrescrito', 'sup'],
      ['Caixa alta', 'span[data-caps]'],
      ['Versalete', 'span[data-small-caps]'],
    ] as const) {
      await session.window.getByRole('button', { name: label, exact: true }).click()
      await expect(editor.locator(selector)).toHaveCount(1)
    }
  })

  test('the paragraph dialog applies spacing and indent', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Parágrafo medido.')

    await menu(session, 'paragraph-setup')
    const dialog = session.window.getByRole('dialog', { name: 'Parágrafo' })
    await expect(dialog).toBeVisible()

    // By role: "Left" is also an alignment option, in the picker's `label`.
    await dialog.getByRole('spinbutton', { name: 'Antes' }).fill('18')
    await dialog.getByRole('spinbutton', { name: 'Esquerda' }).fill('20')
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialog).toBeHidden()

    // A real measure, as pagination will see it.
    const measures = await editor
      .locator('p')
      .first()
      .evaluate((element) => {
        const style = getComputedStyle(element)
        return { before: style.marginTop, indent: style.paddingLeft }
      })

    // 18 pt in CSS pixels, and 20 mm likewise.
    expect(Number.parseFloat(measures.before)).toBeCloseTo(18 * (96 / 72), 0)
    expect(Number.parseFloat(measures.indent)).toBeCloseTo(20 * (96 / 25.4), 0)
  })

  test('1.5 line spacing is 1.5 lines, not CSS 1.5', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Parágrafo de uma linha e meia.')

    await menu(session, 'paragraph-setup')
    const dialog = session.window.getByRole('dialog', { name: 'Parágrafo' })
    await dialog.getByRole('combobox', { name: 'Entrelinha' }).selectOption('1.5')
    await dialog.getByRole('button', { name: 'Aplicar' }).click()

    // Word's multiple is over the natural height of the style font: 1.5 lines in Calibri is
    // `line-height: 1.8311`.
    const ratio = await editor
      .locator('p')
      .first()
      .evaluate((element) => {
        const style = getComputedStyle(element)
        return Number.parseFloat(style.lineHeight) / Number.parseFloat(style.fontSize)
      })

    expect(ratio).toBeCloseTo(1.8311, 2)

    // And the toolbar shows Word's number.
    await expect(session.window.getByRole('combobox', { name: 'Espaçamento entre linhas' })).toHaveValue(
      '1.5',
    )
  })

  test('"Aplicar" without changing anything returns focus to the text', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Sem mudança nenhuma.')

    await menu(session, 'paragraph-setup')
    const dialog = session.window.getByRole('dialog', { name: 'Parágrafo' })
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialog).toBeHidden()

    // "Apply" with nothing to change returns focus to the text, in the frame after closing.
    await expect(editor).toBeFocused()
    await session.window.keyboard.type(' Continua.')
    await expect(editor).toContainText('Sem mudança nenhuma. Continua.')
  })
})
