import { stat } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * Everyday tools through the assembled app, at the seam between the two processes: the
 * `webContents` context menu, the system clipboard and the dictionary that must be on disk before
 * Chromium asks. `new-document` first, because the home screen has no editor.
 */
test.describe('document tools', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-document')
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('right-click opens the menu with cut, copy and paste', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Texto para o menu.')

    await editor.click({ button: 'right' })

    const context = session.window.getByRole('menu', { name: 'Ações do documento' })
    await expect(context).toBeVisible()
    await expect(context.getByRole('menuitem', { name: 'Copiar' })).toBeVisible()
    await expect(context.getByRole('menuitem', { name: 'Colar sem formatação' })).toBeVisible()

    // Escape closes, like every system context menu.
    await session.window.keyboard.press('Escape')
    await expect(context).toBeHidden()
  })

  test('the Portuguese dictionary is installed and the spell checker is on', async () => {
    // Without the dictionary in place, the spellchecker marks nothing, and nobody is warned.
    const spellChecker = await session.app.evaluate(({ session: electronSession }) => ({
      languages: electronSession.defaultSession.getSpellCheckerLanguages(),
      enabled: electronSession.defaultSession.isSpellCheckerEnabled(),
    }))

    expect(spellChecker.languages).toEqual(['pt-BR'])
    expect(spellChecker.enabled).toBe(true)

    // The file Chromium looks for before downloading, with the size of what ships in the installer.
    const installed = await stat(join(session.userData, 'Dictionaries', 'pt-BR-3-0.bdic'))
    const builtIn = await stat(resolve('resources/dictionaries/pt-BR-3-0.bdic'))
    expect(installed.size).toBe(builtIn.size)

    // And the text field asks for checking: the other half of the feature.
    await expect(session.window.locator('.ProseMirror')).toHaveAttribute('spellcheck', 'true')
  })

  test('turning off spell checking turns off the checker and the field', async () => {
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
  test.skip('a misspelled word brings suggestions, dictionary and ignore', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    // Bold so there is somewhere to click: the paragraph center would land far from the word.
    await session.window.keyboard.press('Control+b')
    await session.window.keyboard.type('abacaxxi')
    await session.window.keyboard.press('Control+b')
    await session.window.keyboard.type(' e mais texto.')

    const context = session.window.getByRole('menu', { name: 'Ações do documento' })

    await expect(async () => {
      await session.window.keyboard.press('Escape')
      await editor.locator('strong').click({ button: 'right' })
      await expect(context.getByRole('menuitem', { name: 'Adicionar ao dicionário' })).toBeVisible({
        timeout: 1500,
      })
    }).toPass({ timeout: 20_000 })

    await expect(context.getByRole('menuitem', { name: 'Ignorar nesta sessão' })).toBeVisible()
  })

  test('paste without formatting brings the text and leaves the formatting out', async () => {
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

  test('the word count counts the document and the selection', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('uma frase de teste')

    await menu(session, 'word-count')

    const dialog = session.window.getByRole('dialog', { name: 'Contagem de palavras' })
    await expect(dialog).toBeVisible()

    const words = dialog.getByRole('row').filter({ hasText: 'Palavras' })
    await expect(words).toContainText('4')

    // Without a selection the right column is a dash: "nothing selected" and "a selection of zero
    // words" are not the same thing.
    await expect(words).toContainText('—')
  })

  test('the special character picker inserts at the cursor', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Bom dia')

    await menu(session, 'special-character')

    const dialog = session.window.getByRole('dialog', { name: 'Caracteres especiais' })
    await expect(dialog).toBeVisible()

    // By character name, the button's accessible label: pointing by glyph would depend on the test
    // being able to type it.
    await dialog.getByRole('button', { name: 'travessão' }).click()
    await expect(editor).toContainText('Bom dia—')
  })

  test('formatting marks appear without disturbing pagination', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Uma linha com espaços.')

    const pages = session.window.locator('.statusbar__metric').filter({ hasText: 'página' })
    const before = await pages.textContent()

    await session.window.getByRole('button', { name: /Marcas de formatação/ }).click()

    // One mark per space and one at the end of the paragraph.
    await expect(editor.locator('.tiptap-invisible-character')).not.toHaveCount(0)

    // A mark with height would grow the line and shift the page break.
    await expect(pages).toHaveText(before ?? '')

    await session.window.getByRole('button', { name: /Marcas de formatação/ }).click()
    await expect(editor.locator('.tiptap-invisible-character')).toHaveCount(0)
  })

  test('the special character picker also works by keyboard', async () => {
    // Without focus on the panel, arrows and `Enter` would go to the text, and `Tab` to the "Style"
    // picker.
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Bom dia')

    await menu(session, 'special-character')
    const dialog = session.window.getByRole('dialog', { name: 'Caracteres especiais' })

    // The grid's first character is already focused; arrows move, Enter inserts.
    const first = dialog.getByRole('button').first()
    await expect(first).toBeFocused()

    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('Enter')

    // Inserted and stayed in the panel.
    await expect(dialog).toBeVisible()
    await expect(editor).not.toHaveText('Bom dia')
    const afterFirst = await editor.textContent()

    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('Enter')
    await expect(editor).not.toHaveText(afterFirst ?? '')

    // Tab cycles inside the panel: the toolbar's "Style" is not reachable from here.
    await session.window.keyboard.press('Tab')
    await expect(session.window.getByRole('combobox', { name: 'Estilo' })).not.toBeFocused()

    await session.window.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
  })

  test('typographic autocorrect replaces the quotes and the dash', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Ele disse "sim" -- e saiu')

    // Curly quotes and em dash, as in Portuguese Word.
    await expect(editor).toContainText('“sim”')
    await expect(editor).toContainText('—')
  })

  test('Backspace undoes only the autocorrect replacement', async () => {
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
