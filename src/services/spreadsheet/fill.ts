/** Dragging copies the shifted formula (`=B2*C2` → `=B3*C3`), not the result. */

import { normalizeRange, type Range } from './edit.js'
import { translateFormula } from './formula/adjust.js'
import { getCell, setCell, type Cell, type Sheet } from './model.js'

/**
 * `target` includes the source, which is not rewritten. Repetition is cyclic, as in Excel: two rows
 * dragged over six repeat the pattern three times.
 */
export function fillRange(sheet: Sheet, source: Range, target: Range): Sheet {
  const from = normalizeRange(source)
  const to = normalizeRange(target)

  const rows = from.toRow - from.fromRow + 1
  const columns = from.toColumn - from.fromColumn + 1
  if (rows <= 0 || columns <= 0) return sheet

  let updated = sheet

  for (let row = to.fromRow; row <= to.toRow; row++) {
    for (let column = to.fromColumn; column <= to.toColumn; column++) {
      if (row >= from.fromRow && row <= from.toRow && column >= from.fromColumn && column <= from.toColumn) {
        continue
      }

      // Dragging upward gives a negative difference, and `%` in JavaScript returns it negative.
      const sourceRow = from.fromRow + ((((row - from.fromRow) % rows) + rows) % rows)
      const sourceColumn = from.fromColumn + ((((column - from.fromColumn) % columns) + columns) % columns)

      updated = setCell(
        updated,
        row,
        column,
        copyOf(getCell(updated, sourceRow, sourceColumn), row - sourceRow, column - sourceColumn),
      )
    }
  }

  return updated
}

/** The style comes along, as in Excel; the value does not, because recalculation fills it. */
function copyOf(source: Cell | undefined, rowDelta: number, columnDelta: number): Cell {
  if (source === undefined) return {}

  const cell: Cell =
    source.formula === undefined
      ? { ...(source.value === undefined ? {} : { value: source.value }) }
      : { formula: translateFormula(source.formula, rowDelta, columnDelta) }

  return source.style === undefined ? cell : { ...cell, style: source.style }
}
