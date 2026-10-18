import { expect, type Page } from '@playwright/test'

/**
 * A grid cell, by zero-based coordinates of the sheet. The ARIA indexes, not `data-rgcol`, which
 * counts from the start of each frozen or scrolling pane.
 */
export function cell(window: Page, row: number, column: number) {
  return window
    .locator(`revogr-data [role="gridcell"][aria-rowindex="${row}"][aria-colindex="${column}"]`)
    .first()
}

/** The grid only draws cells after mounting; clicking earlier selects nothing. */
export async function gridReady(window: Page): Promise<void> {
  await expect(window.locator('revo-grid')).toBeVisible()
  await expect(cell(window, 0, 0)).toBeVisible()
}

export async function select(window: Page, row: number, column: number): Promise<void> {
  // The click repeats until the selection moves: the grid redraws after opening.
  await expect(async () => {
    await cell(window, row, column).click()
    await expect(window.locator('.formula-bar__ref')).toHaveText(reference(row, column), {
      timeout: 1000,
    })
  }).toPass({ timeout: 15_000 })
}

/** Selects the cell and writes into it through the formula bar. */
export async function write(window: Page, row: number, column: number, text: string): Promise<void> {
  await select(window, row, column)
  const input = window.locator('.formula-bar__input')
  await input.fill(text)
  await input.press('Enter')
}

export function reference(row: number, column: number): string {
  return `${String.fromCharCode(65 + column)}${row + 1}`
}

/** From the first cell down the column, with Shift and the arrow keys, as in any spreadsheet. */
export async function selectDown(window: Page, row: number, column: number, rows: number): Promise<void> {
  const range = `${reference(row, column)}:${reference(row + rows, column)}`
  // Repeated as `select`: a key pressed while the grid still redraws is lost.
  await expect(async () => {
    await select(window, row, column)
    for (let step = 0; step < rows; step++) await window.keyboard.press('Shift+ArrowDown')
    await expect(window.locator('.formula-bar__ref')).toHaveText(range, { timeout: 1000 })
  }).toPass({ timeout: 15_000 })
}

/** Drags the small square at the selection corner, starting a few pixels inside. */
export async function dragHandle(
  window: Page,
  fromRow: number,
  fromColumn: number,
  toRow: number,
  toColumn: number,
): Promise<void> {
  await select(window, fromRow, fromColumn)

  const origin = await cell(window, fromRow, fromColumn).boundingBox()
  const destination = await cell(window, toRow, toColumn).boundingBox()
  if (origin === null || destination === null) throw new Error('célula sem posição na tela')

  await window.mouse.move(origin.x + origin.width - 2, origin.y + origin.height - 2)
  await window.mouse.down()
  await window.mouse.move(destination.x + destination.width / 2, destination.y + destination.height / 2, {
    steps: 8,
  })
  await window.mouse.up()
}
