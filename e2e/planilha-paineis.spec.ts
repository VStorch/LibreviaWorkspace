import { expect, test, type Page } from '@playwright/test'
import { launch, menu, type Session } from './app.js'
import { cell, gridReady, select, write } from './sheet-grid.js'

/** Columns A and B and rows 1 and 2 frozen. */
async function freeze(window: Page): Promise<void> {
  await select(window, 2, 2)
  await window.getByRole('button', { name: 'Congelar até a seleção' }).click()
}

async function selectAcross(window: Page, to: { row: number; column: number }, text: string): Promise<void> {
  await expect(async () => {
    await select(window, 0, 0)
    await cell(window, to.row, to.column).click({ modifiers: ['Shift'] })
    await expect(window.locator('.formula-bar__ref')).toHaveText(text, { timeout: 1000 })
  }).toPass({ timeout: 15_000 })
}

async function drag(window: Page, to: { row: number; column: number }): Promise<void> {
  const box = await cell(window, to.row, to.column).boundingBox()
  await window.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2, { steps: 10 })
}

test.describe('painéis congelados na planilha', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-spreadsheet')
    await gridReady(session.window)
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('arrastar com o mouse estende a seleção através da divisa e volta', async () => {
    const { window } = session
    const ref = window.locator('.formula-bar__ref')
    await freeze(window)
    await select(window, 0, 0)
    const start = await cell(window, 0, 0).boundingBox()
    await window.mouse.move(start!.x + 5, start!.y + 5)
    await window.mouse.down()

    await drag(window, { row: 3, column: 3 })
    await expect(ref).toHaveText('A1:D4')
    await expect(window.locator('revogr-overlay-selection .selection-border-range')).toHaveCount(2)
    await drag(window, { row: 1, column: 1 })
    await expect(ref).toHaveText('A1:B2')
    await window.mouse.up()

    await expect(ref).toHaveText('A1:B2')
    await expect(window.locator('revogr-overlay-selection .selection-border-range')).toHaveCount(1)
  })

  test('Delete apaga a seleção inteira, dos dois lados da divisa', async () => {
    const { window } = session
    for (const [column, text] of ['a', 'b', 'c', 'd'].entries()) await write(window, 0, column, text)
    await freeze(window)
    await selectAcross(window, { row: 0, column: 2 }, 'A1:C1')

    await window.keyboard.press('Delete')
    for (const column of [0, 1, 2]) await expect(cell(window, 0, column)).toHaveText('')
    await expect(cell(window, 0, 3)).toHaveText('d')
  })

  test('copiar, colar e recortar levam a seleção inteira, em qualquer painel', async () => {
    const { window, app } = session
    await write(window, 0, 0, 'a')
    await write(window, 0, 1, 'b')
    await write(window, 0, 2, '=1+1')
    await freeze(window)
    await selectAcross(window, { row: 0, column: 2 }, 'A1:C1')

    await window.keyboard.press('Control+c')
    await expect.poll(() => app.evaluate(({ clipboard }) => clipboard.readText())).toBe('a\tb\t2')

    await select(window, 5, 0)
    await window.keyboard.press('Control+v')
    await expect(cell(window, 5, 0)).toHaveText('a')
    await expect(cell(window, 5, 2)).toHaveText('2')

    await app.evaluate(({ clipboard }) => clipboard.writeText('x\ty\nz\tw'))
    await select(window, 6, 3)
    await window.keyboard.press('Control+v')
    await expect(cell(window, 6, 4)).toHaveText('y')
    await expect(cell(window, 7, 3)).toHaveText('z')
    await window.keyboard.press('Control+z')
    await expect(cell(window, 6, 3)).toHaveText('')
    await expect(cell(window, 7, 4)).toHaveText('')

    await window.locator('.formula-bar__ref').selectText()
    await selectAcross(window, { row: 0, column: 2 }, 'A1:C1')
    await window.keyboard.press('Control+x')
    await expect.poll(() => app.evaluate(({ clipboard }) => clipboard.readText())).toBe('a\tb\t2')
    for (const column of [0, 1, 2]) await expect(cell(window, 0, column)).toHaveText('')
  })
})
