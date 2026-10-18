import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'
import { cell, dragHandle, select, write } from './sheet-grid.js'

/**
 * The fill handle carries the formula, not the result: drag with the mouse, because a unit test
 * cannot know whether the grid event was subscribed.
 */
test.describe('alça de preenchimento', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('arrastar uma fórmula desloca as referências', async () => {
    await menu(session, 'new-spreadsheet')
    await expect(cell(session.window, 0, 0)).toBeVisible()

    await write(session.window, 0, 0, '2')
    await write(session.window, 1, 0, '3')
    await write(session.window, 0, 1, '10')
    await write(session.window, 1, 1, '20')
    await write(session.window, 0, 2, '=A1*B1')

    await expect(cell(session.window, 0, 2)).toHaveText('20')

    await dragHandle(session.window, 0, 2, 1, 2)

    // Copying the displayed text would give 20 here. The shifted formula gives 60.
    await expect(cell(session.window, 1, 2)).toHaveText('60')

    await select(session.window, 1, 2)
    await expect(session.window.locator('.formula-bar__input')).toHaveValue('=A2*B2')
  })
})
