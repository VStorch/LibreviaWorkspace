import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'
import { cell } from './sheet-grid.js'

/**
 * Entering a column of numbers nonstop: the grid waits 70 ms after Enter to move focus down, and a
 * key arriving in that window would go to the previous cell (`1200` below `980` becomes `200`). It
 * lies where the grid's clock meets the typist's, which a unit test cannot reach. The formula bar
 * does not go through it.
 */
test.describe('continuous typing in the spreadsheet', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('a column typed without pause arrives whole', async () => {
    await menu(session, 'new-spreadsheet')
    await expect(cell(session.window, 0, 0)).toBeVisible()

    const values = ['980', '1200', '2450', '860']
    await cell(session.window, 0, 0).click()

    for (const value of values) {
      // 60 ms per key is fast keypad typing, with no pause between Enter and the next number, which
      // is how a column is entered.
      await session.window.keyboard.type(value, { delay: 60 })
      await session.window.keyboard.press('Enter')
    }

    for (const [line, value] of values.entries()) {
      await expect(cell(session.window, line, 0)).toHaveText(value)
    }
  })
})
