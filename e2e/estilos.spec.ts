/// <reference lib="dom" />
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithNamedStyles, entryOf } from './fixtures.js'

/**
 * The styles pane: view, apply, modify, create and clear, with the screen changing, undo, and the
 * modified style going to the `.docx` and back.
 */
test.describe('painel de estilos', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-estilos-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('o documento novo mostra os estilos embutidos e o do cursor', async () => {
    await menu(session, 'new-document')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })

    // A new document paragraph points to no style: the default applies, and the pane names it.
    await expect(panel).toContainText('Parágrafo do cursor: Normal')
    await expect(panel).toContainText('Título 1')

    // Becoming a heading changes the answer live, with the pane open.
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption({ label: 'Título 1' })
    await expect(panel).toContainText('Parágrafo do cursor: Título 1')
  })

  test('mostra os estilos que o .docx traz, com o nome que o autor deu', async () => {
    const origem = join(folder, 'estilos.docx')
    await writeFile(origem, await docxWithNamedStyles())

    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Um trecho citado.')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })

    // The author's style shows by name, and the heading by its translated internal name; the
    // `Ttulo1` id belongs to the file.
    await expect(panel).toContainText('Citação recuada')
    await expect(panel).toContainText('Título 1')

    // Word's machinery stays out of the list, through `w:semiHidden`.
    await expect(panel).not.toContainText('Default Paragraph Font')

    // The cursor opens on the heading, whose style comes from the file with a Portuguese id.
    await expect(panel).toContainText('Parágrafo do cursor: Título 1')

    // The filter makes a Word document's list readable, since it declares dozens.
    await panel.getByRole('combobox', { name: 'Mostrar' }).selectOption('character')
    await expect(panel).not.toContainText('Citação recuada')
  })

  test('aplicar pelo painel troca o bloco, e desfazer o devolve', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Relatório')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel.locator('.styles-list__item', { hasText: 'Título 1' }).click()
    await panel.getByRole('button', { name: 'Aplicar' }).click()

    await expect(editor.locator('h1')).toHaveText('Relatório')
    await expect(panel).toContainText('Parágrafo do cursor: Título 1')

    await session.window.keyboard.press('Control+z')
    await expect(editor.locator('h1')).toHaveCount(0)
    await expect(editor.locator('p').first()).toHaveText('Relatório')
  })

  test('modificar o estilo muda a tela na hora', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Corpo do texto.')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel
      .locator('.styles-list__item')
      .filter({ has: session.window.locator('.styles-list__name', { hasText: /^Normal$/ }) })
      .click()
    await panel.getByRole('button', { name: 'Modificar…' }).click()

    // A builtin style is not renamed: Word recognizes it by its name.
    await expect(panel.getByRole('textbox', { name: 'Nome' })).toBeDisabled()
    await panel.getByRole('spinbutton', { name: 'Tamanho (pt)' }).fill('20')
    await panel.getByRole('button', { name: 'OK' }).click()

    // 20 pt in CSS pixels: the style rule was regenerated.
    const size = await editor
      .locator('p')
      .first()
      .evaluate((element) => getComputedStyle(element).fontSize)
    expect(Number.parseFloat(size)).toBeCloseTo(20 * (96 / 72), 0)
    await expect(session.window.locator('.statusbar__state')).not.toHaveText('Salvo')
  })

  test('Enter no fim do título abre o estilo seguinte', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Capítulo')
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption({ label: 'Título 1' })

    await editor.locator('h1').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Texto depois do título.')

    await expect(editor.locator('p').first()).toHaveText('Texto depois do título.')
    await expect(session.window.getByRole('combobox', { name: 'Estilo' })).toHaveValue('Normal')
  })

  test('limpar a formatação tira o direto e deixa o estilo', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.press('Control+b')
    await session.window.keyboard.type('Negrito à mão')
    await expect(editor.locator('strong')).toHaveCount(1)

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel.getByRole('button', { name: 'Limpar formatação' }).click()

    await expect(editor.locator('strong')).toHaveCount(0)
    await expect(editor.locator('p').first()).toHaveText('Negrito à mão')
  })

  test('o estilo modificado vai para o .docx e volta dele', async () => {
    const origem = join(folder, 'estilos.docx')
    await writeFile(origem, await docxWithNamedStyles())

    await stubDialogs(session.app, { open: origem, save: origem, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('Um trecho citado.')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel.locator('.styles-list__item', { hasText: 'Citação recuada' }).click()
    await panel.getByRole('button', { name: 'Modificar…' }).click()
    await panel.getByRole('spinbutton', { name: 'Tamanho (pt)' }).fill('18')
    await panel.getByRole('button', { name: 'OK' }).click()
    await panel.getByRole('button', { name: 'Fechar' }).click()

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // Only the style changed in the file: the paragraph still only points to it.
    expect(await entryOf(origem, 'word/styles.xml')).toContain('<w:sz w:val="36"')
    expect(await entryOf(origem, 'word/document.xml')).not.toContain('w:sz')

    await menu(session, 'open')
    const quote = editor.locator('[data-style-id="Citao"]')
    await expect(quote).toHaveText('Um trecho citado.')
    // With waiting: the measure could land on the previous document's paragraph.
    await expect
      .poll(async () =>
        Number.parseFloat(await quote.evaluate((element) => getComputedStyle(element).fontSize)),
      )
      .toBeCloseTo(18 * (96 / 72), 0)
  })

  test('o texto importado segue o estilo: modificado, e de título a Normal', async () => {
    // The run, not the paragraph: the style font stuck in a mark would change one and not the
    // other.
    const origem = join(folder, 'estilos.docx')
    await writeFile(origem, await docxWithNamedStyles())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('Um trecho citado.')

    const look = (text: string) =>
      editor.evaluate((root, needle) => {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
        for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
          if (node.textContent?.includes(needle) !== true) continue
          const style = getComputedStyle(node.parentElement!)
          return { size: Number.parseFloat(style.fontSize), weight: style.fontWeight }
        }
        return null
      }, text)

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel.locator('.styles-list__item', { hasText: 'Citação recuada' }).click()
    await panel.getByRole('button', { name: 'Modificar…' }).click()
    await panel.getByRole('spinbutton', { name: 'Tamanho (pt)' }).fill('18')
    await panel.getByRole('button', { name: 'OK' }).click()
    await panel.getByRole('button', { name: 'Fechar' }).click()

    expect((await look('Um trecho citado.'))?.size).toBeCloseTo(18 * (96 / 72), 0)

    // Switched to Normal, the heading goes back to 11 pt without bold: nothing stuck in the run.
    expect((await look('Relatório anual'))?.weight).toBe('700')
    await editor.getByText('Relatório anual').click()
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption({ label: 'Normal' })
    const normal = await look('Relatório anual')
    expect(normal?.size).toBeCloseTo(11 * (96 / 72), 0)
    expect(normal?.weight).toBe('400')
  })

  test('o foco circula dentro do painel, e Escape continua fechando', async () => {
    await menu(session, 'new-document')
    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()

    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await expect(panel.getByRole('button', { name: 'Fechar' })).toBeFocused()

    // Tab cycles in the pane, or it would land in the text, where Escape does not close it.
    await panel.press('Tab')
    await expect(panel.getByRole('combobox', { name: 'Mostrar' })).toBeFocused()

    await panel.press('Shift+Tab')
    await expect(panel.getByRole('button', { name: 'Fechar' })).toBeFocused()

    await panel.press('Escape')
    await expect(panel).toBeHidden()
  })
})
