import { stat } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * Everyday tools through the assembled app, at the seam between the two processes: the
 * `webContents` context menu, the system clipboard and the dictionary that must be on disk before
 * Chromium asks. `new-document` first, because the home screen has no editor.
 */
test.describe('ferramentas do documento', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-document')
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('o botão direito abre o menu com recortar, copiar e colar', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Texto para o menu.')

    await editor.click({ button: 'right' })

    const contexto = session.window.getByRole('menu', { name: 'Ações do documento' })
    await expect(contexto).toBeVisible()
    await expect(contexto.getByRole('menuitem', { name: 'Copiar' })).toBeVisible()
    await expect(contexto.getByRole('menuitem', { name: 'Colar sem formatação' })).toBeVisible()

    // Escape closes, like every system context menu.
    await session.window.keyboard.press('Escape')
    await expect(contexto).toBeHidden()
  })

  test('o dicionário de português está instalado e o corretor, ligado', async () => {
    // Without the dictionary in place, the spellchecker marks nothing, and nobody is warned.
    const corretor = await session.app.evaluate(({ session: electronSession }) => ({
      idiomas: electronSession.defaultSession.getSpellCheckerLanguages(),
      ligado: electronSession.defaultSession.isSpellCheckerEnabled(),
    }))

    expect(corretor.idiomas).toEqual(['pt-BR'])
    expect(corretor.ligado).toBe(true)

    // The file Chromium looks for before downloading, with the size of what ships in the installer.
    const instalado = await stat(join(session.userData, 'Dictionaries', 'pt-BR-3-0.bdic'))
    const embutido = await stat(resolve('resources/dictionaries/pt-BR-3-0.bdic'))
    expect(instalado.size).toBe(embutido.size)

    // And the text field asks for checking: the other half of the feature.
    await expect(session.window.locator('.ProseMirror')).toHaveAttribute('spellcheck', 'true')
  })

  test('desligar a verificação ortográfica desliga o corretor e o campo', async () => {
    // Through the native menu item's path, which no test clicks. The `window.api` type lives in the
    // renderer declaration, outside this project: hence the cast.
    await session.window.evaluate(async () => {
      const api = (
        window as unknown as { api: { preferences: { set: (patch: unknown) => Promise<unknown> } } }
      ).api
      await api.preferences.set({ spellcheck: false })
    })

    await expect(session.window.locator('.ProseMirror')).toHaveAttribute('spellcheck', 'false')
    expect(
      await session.app.evaluate(({ session: electronSession }) =>
        electronSession.defaultSession.isSpellCheckerEnabled(),
      ),
    ).toBe(false)
  })

  // Chromium only marks a misspelled word with the frame truly focused, which automated runs do not
  // guarantee. The path to suggestions is covered above.
  test.skip('a palavra errada traz sugestão, dicionário e ignorar', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    // Bold so there is somewhere to click: the paragraph center would land far from the word.
    await session.window.keyboard.press('Control+b')
    await session.window.keyboard.type('abacaxxi')
    await session.window.keyboard.press('Control+b')
    await session.window.keyboard.type(' e mais texto.')

    const contexto = session.window.getByRole('menu', { name: 'Ações do documento' })

    await expect(async () => {
      await session.window.keyboard.press('Escape')
      await editor.locator('strong').click({ button: 'right' })
      await expect(contexto.getByRole('menuitem', { name: 'Adicionar ao dicionário' })).toBeVisible({
        timeout: 1500,
      })
    }).toPass({ timeout: 20_000 })

    await expect(contexto.getByRole('menuitem', { name: 'Ignorar nesta sessão' })).toBeVisible()
  })

  test('colar sem formatação traz o texto e deixa a formatação de fora', async () => {
    await session.app.evaluate(({ clipboard }) => {
      clipboard.write({
        text: 'primeira linha\nsegunda linha',
        html: '<p><strong>primeira linha</strong></p><p><em>segunda linha</em></p>',
      })
    })

    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await menu(session, 'paste-without-format')

    await expect(editor).toContainText('primeira linha')
    await expect(editor).toContainText('segunda linha')
    // Two lines become two paragraphs, and neither carries the bold or italic from the clipboard
    // HTML.
    await expect(editor.locator('strong')).toHaveCount(0)
    await expect(editor.locator('em')).toHaveCount(0)
  })

  test('a contagem de palavras conta o documento e a seleção', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('uma frase de teste')

    await menu(session, 'word-count')

    const dialogo = session.window.getByRole('dialog', { name: 'Contagem de palavras' })
    await expect(dialogo).toBeVisible()

    const palavras = dialogo.getByRole('row').filter({ hasText: 'Palavras' })
    await expect(palavras).toContainText('4')

    // Without a selection the right column is a dash: "nothing selected" and "a selection of zero
    // words" are not the same thing.
    await expect(palavras).toContainText('—')
  })

  test('o seletor de caracteres especiais insere no cursor', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Bom dia')

    await menu(session, 'special-character')

    const dialogo = session.window.getByRole('dialog', { name: 'Caracteres especiais' })
    await expect(dialogo).toBeVisible()

    // By character name, the button's accessible label: pointing by glyph would depend on the test
    // being able to type it.
    await dialogo.getByRole('button', { name: 'travessão' }).click()
    await expect(editor).toContainText('Bom dia—')
  })

  test('as marcas de formatação aparecem sem mexer na paginação', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Uma linha com espaços.')

    const paginas = session.window.locator('.statusbar__metric').filter({ hasText: 'página' })
    const antes = await paginas.textContent()

    await session.window.getByRole('button', { name: /Marcas de formatação/ }).click()

    // One mark per space and one at the end of the paragraph.
    await expect(editor.locator('.tiptap-invisible-character')).not.toHaveCount(0)

    // A mark with height would grow the line and shift the page break.
    await expect(paginas).toHaveText(antes ?? '')

    await session.window.getByRole('button', { name: /Marcas de formatação/ }).click()
    await expect(editor.locator('.tiptap-invisible-character')).toHaveCount(0)
  })

  test('o seletor de caracteres especiais também anda pelo teclado', async () => {
    // Without focus on the panel, arrows and `Enter` would go to the text, and `Tab` to the "Style"
    // picker.
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Bom dia')

    await menu(session, 'special-character')
    const dialogo = session.window.getByRole('dialog', { name: 'Caracteres especiais' })

    // The grid's first character is already focused; arrows move, Enter inserts.
    const primeiro = dialogo.getByRole('button').first()
    await expect(primeiro).toBeFocused()

    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('Enter')

    // Inserted and stayed in the panel.
    await expect(dialogo).toBeVisible()
    await expect(editor).not.toHaveText('Bom dia')
    const depoisDoPrimeiro = await editor.textContent()

    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('Enter')
    await expect(editor).not.toHaveText(depoisDoPrimeiro ?? '')

    // Tab cycles inside the panel: the toolbar's "Style" is not reachable from here.
    await session.window.keyboard.press('Tab')
    await expect(session.window.getByRole('combobox', { name: 'Estilo' })).not.toBeFocused()

    await session.window.keyboard.press('Escape')
    await expect(dialogo).toBeHidden()
  })

  test('a autocorreção tipográfica troca as aspas e o travessão', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Ele disse "sim" -- e saiu')

    // Curly quotes and em dash, as in Portuguese Word.
    await expect(editor).toContainText('“sim”')
    await expect(editor).toContainText('—')
  })

  test('o Backspace desfaz só a substituição da autocorreção', async () => {
    // `--silent` goes back to `--silent`: Tiptap's `Keymap` handles the key, and the `undoable` the
    // autocorrect wrapper keeps preserves it (`editor-extensions.ts`).
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('rodar --')
    await expect(editor).toContainText('rodar —')

    // Right after the replacement, as in Word: the plugin keeps only the last correction.
    await session.window.keyboard.press('Backspace')
    await session.window.keyboard.type('silent')
    await expect(editor).toContainText('rodar --silent')
  })
})
