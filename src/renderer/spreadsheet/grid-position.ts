/**
 * The grid is a web component: the position is in the attributes it leaves on the cell DOM, its
 * public contract. The ARIA indexes, not `data-rgrow`/`data-rgcol`: those count from the start of
 * each frozen or scrolling pane, and the row numbers on the left carry them too.
 */

const CELL = '[role="gridcell"]'

function indexOf(cell: Element, name: string): number | null {
  const index = Number.parseInt(cell.getAttribute(name) ?? '', 10)
  return Number.isInteger(index) ? index : null
}

/**
 * `null` on the headers and the empty area. `composedPath()[0]`, not `event.target`, which with
 * shadow DOM would arrive retargeted to the host.
 */
export function gridPositionOf(event: Event): { row: number; column: number } | null {
  const deepest = event.composedPath()[0]
  const cell = deepest instanceof Element ? deepest.closest(CELL) : null
  if (cell === null) return null

  const row = indexOf(cell, 'aria-rowindex')
  const column = indexOf(cell, 'aria-colindex')
  return row === null || column === null ? null : { row, column }
}

/** The grid's names for the panes; only columns freeze into a pane of their own. */
export const FROZEN_COLUMNS_PANE = 'colPinStart'

/** Rows freeze in place, not into a pane: every range of the sheet lives in the scrolling rows. */
export const SHEET_ROWS_PANE = 'rgRow'

/** From a pane's own count to the sheet's: the scrolling pane starts after the frozen columns. */
export function sheetColumnOf(paneColumn: number, pane: string | undefined, frozenColumns: number): number {
  return pane === FROZEN_COLUMNS_PANE ? paneColumn : paneColumn + frozenColumns
}
