import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'
import { cell, gridReady, select } from './sheet-grid.js'

/**
 * Through the real key, not the menu command: on Windows and Linux Chromium takes `Ctrl+Shift+V` in
 * an editable field as its own "paste and match style", and the menu accelerator never fires. The
 * empty line is the sign of the app's paste: Chromium's, through ProseMirror, drops it.
 */
test.describe('paste without formatting', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('Ctrl+Shift+V in the document pastes the text with the app rules, empty lines included', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.app.evaluate(({ clipboard }) => {
      clipboard.write({ text: 'um\n\ndois', html: '<p><b>um</b></p><p></p><p><i>dois</i></p>' })
    })

    await session.window.keyboard.press('Control+Shift+V')

    await expect(editor.locator('p')).toHaveText(['um', '', 'dois'])
    await expect(editor.locator('strong, em')).toHaveCount(0)
  })

  test('Ctrl+Shift+V and the menu item paste the text into the spreadsheet', async () => {
    const { window, app } = session
    await menu(session, 'new-spreadsheet')
    await gridReady(window)
    await app.evaluate(({ clipboard }) => {
      clipboard.write({ text: 'a\tb\nc\td', html: '<table><tr><td><b>a</b></td></tr></table>' })
    })

    await select(window, 0, 0)
    await window.keyboard.press('Control+Shift+V')
    await expect(cell(window, 0, 1)).toHaveText('b')
    await expect(cell(window, 1, 0)).toHaveText('c')

    await select(window, 4, 2)
    await menu(session, 'paste-without-format')
    await expect(cell(window, 4, 3)).toHaveText('b')
    await expect(cell(window, 5, 2)).toHaveText('c')
  })
})
