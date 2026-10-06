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
test.describe('formatação do documento', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-formatacao-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('o sobrescrito do documento aparece na tela e volta para o arquivo', async () => {
    const origem = join(folder, 'formula.docx')
    const destino = join(folder, 'salva.docx')
    await writeFile(origem, await docxWithVerticalAlignment())

    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
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

    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo).toContain('w:vertAlign w:val="superscript"')
    expect(corpo).toContain('w:vertAlign w:val="subscript"')

    // Nothing in the notice: the writer knows how to write this formatting.
    await expect(session.window.locator('.banner--notice')).toHaveCount(0)
  })

  test('os botões de sobrescrito, caixa alta e versalete marcam o texto', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('marcado')
    await session.window.keyboard.press('Control+a')

    for (const [rotulo, seletor] of [
      ['Sobrescrito', 'sup'],
      ['Caixa alta', 'span[data-caps]'],
      ['Versalete', 'span[data-small-caps]'],
    ] as const) {
      await session.window.getByRole('button', { name: rotulo, exact: true }).click()
      await expect(editor.locator(seletor)).toHaveCount(1)
    }
  })

  test('o diálogo de parágrafo aplica espaçamento e recuo', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Parágrafo medido.')

    await menu(session, 'paragraph-setup')
    const dialogo = session.window.getByRole('dialog', { name: 'Parágrafo' })
    await expect(dialogo).toBeVisible()

    // By role: "Left" is also an alignment option, in the picker's `label`.
    await dialogo.getByRole('spinbutton', { name: 'Antes' }).fill('18')
    await dialogo.getByRole('spinbutton', { name: 'Esquerda' }).fill('20')
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialogo).toBeHidden()

    // A real measure, as pagination will see it.
    const medidas = await editor
      .locator('p')
      .first()
      .evaluate((elemento) => {
        const estilo = getComputedStyle(elemento)
        return { antes: estilo.marginTop, recuo: estilo.paddingLeft }
      })

    // 18 pt in CSS pixels, and 20 mm likewise.
    expect(Number.parseFloat(medidas.antes)).toBeCloseTo(18 * (96 / 72), 0)
    expect(Number.parseFloat(medidas.recuo)).toBeCloseTo(20 * (96 / 25.4), 0)
  })

  test('a entrelinha de 1,5 linha é 1,5 linha, e não 1,5 de CSS', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Parágrafo de uma linha e meia.')

    await menu(session, 'paragraph-setup')
    const dialogo = session.window.getByRole('dialog', { name: 'Parágrafo' })
    await dialogo.getByRole('combobox', { name: 'Entrelinha' }).selectOption('1.5')
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()

    // Word's multiple is over the natural height of the style font: 1.5 lines in Calibri is
    // `line-height: 1.8311`.
    const proporcao = await editor
      .locator('p')
      .first()
      .evaluate((elemento) => {
        const estilo = getComputedStyle(elemento)
        return Number.parseFloat(estilo.lineHeight) / Number.parseFloat(estilo.fontSize)
      })

    expect(proporcao).toBeCloseTo(1.8311, 2)

    // And the toolbar shows Word's number.
    await expect(session.window.getByRole('combobox', { name: 'Espaçamento entre linhas' })).toHaveValue(
      '1.5',
    )
  })

  test('"Aplicar" sem mexer em nada devolve o foco ao texto', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Sem mudança nenhuma.')

    await menu(session, 'paragraph-setup')
    const dialogo = session.window.getByRole('dialog', { name: 'Parágrafo' })
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialogo).toBeHidden()

    // "Apply" with nothing to change returns focus to the text, in the frame after closing.
    await expect(editor).toBeFocused()
    await session.window.keyboard.type(' Continua.')
    await expect(editor).toContainText('Sem mudança nenhuma. Continua.')
  })
})
