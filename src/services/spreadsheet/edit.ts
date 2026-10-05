import { parseInput } from './format.js'
import { cellRef, getCell, setCell, type BorderSide, type Cell, type CellStyle, type Sheet } from './model.js'

/** Zero-based and inclusive. */
export interface Range {
  readonly fromRow: number
  readonly fromColumn: number
  readonly toRow: number
  readonly toColumn: number
}

export function normalizeRange(range: Range): Range {
  return {
    fromRow: Math.min(range.fromRow, range.toRow),
    toRow: Math.max(range.fromRow, range.toRow),
    fromColumn: Math.min(range.fromColumn, range.toColumn),
    toColumn: Math.max(range.fromColumn, range.toColumn),
  }
}

export function singleCell(row: number, column: number): Range {
  return { fromRow: row, toRow: row, fromColumn: column, toColumn: column }
}

/** "B3" for a cell, "B3:D9" for a range. */
export function describeRange(range: Range): string {
  const { fromRow, fromColumn, toRow, toColumn } = normalizeRange(range)
  const start = cellRef(fromRow, fromColumn)
  return fromRow === toRow && fromColumn === toColumn ? start : `${start}:${cellRef(toRow, toColumn)}`
}

export function rangeContains(range: Range, row: number, column: number): boolean {
  const { fromRow, fromColumn, toRow, toColumn } = normalizeRange(range)
  return row >= fromRow && row <= toRow && column >= fromColumn && column <= toColumn
}

export function* cellsIn(range: Range): Generator<{ row: number; column: number }> {
  const { fromRow, fromColumn, toRow, toColumn } = normalizeRange(range)
  for (let row = fromRow; row <= toRow; row++) {
    for (let column = fromColumn; column <= toColumn; column++) {
      yield { row, column }
    }
  }
}

/** Merges with the existing style: bold does not erase the background. */
export function applyStyle(sheet: Sheet, range: Range, change: Partial<CellStyle>): Sheet {
  let updated = sheet

  for (const { row, column } of cellsIn(range)) {
    const cell = getCell(updated, row, column) ?? {}
    const style = clean({ ...cell.style, ...change })

    updated = setCell(updated, row, column, toCell(cell, style))
  }

  return updated
}

/** If everything is already on, turns it off; otherwise turns everything on. */
export function toggleStyle(sheet: Sheet, range: Range, key: 'bold' | 'italic' | 'underline'): Sheet {
  const allOn = [...cellsIn(range)].every(
    ({ row, column }) => getCell(sheet, row, column)?.style?.[key] === true,
  )

  return applyStyle(sheet, range, { [key]: allOn ? undefined : true })
}

/** Empty `sides` removes the borders; every cell gets the same sides. */
export function applyBorders(sheet: Sheet, range: Range, sides: readonly BorderSide[]): Sheet {
  return applyStyle(sheet, range, { borders: sides.length === 0 ? undefined : [...sides] })
}

/** Preserves the formatting, as Delete does. */
export function clearContents(sheet: Sheet, range: Range): Sheet {
  let updated = sheet

  for (const { row, column } of cellsIn(range)) {
    const cell = getCell(updated, row, column)
    if (cell === undefined) continue

    updated = setCell(updated, row, column, cell.style === undefined ? {} : { style: cell.style })
  }

  return updated
}

/**
 * A leading `=` tells a formula from text. The value is left to recalculation, which knows the
 * order.
 */
export function writeText(sheet: Sheet, row: number, column: number, text: string): Sheet {
  const previous = getCell(sheet, row, column)

  // A format recognized while typing does not erase the one chosen by hand.
  const keepStyle = (cell: Cell, fallback?: Partial<CellStyle>): Cell => {
    const style = previous?.style ?? fallback
    return style === undefined ? cell : { ...cell, style }
  }

  if (text.startsWith('=')) return setCell(sheet, row, column, keepStyle({ formula: text }))

  const parsed = parseInput(text)
  return setCell(sheet, row, column, keepStyle({ value: parsed.value }, parsed.style))
}

/**
 * The sheet grows along: otherwise the last row would leave the visible area and stay in the file.
 */
export function insertRows(sheet: Sheet, at: number, count = 1): Sheet {
  return count <= 0 ? sheet : shiftRows(sheet, at, count)
}

export function deleteRows(sheet: Sheet, at: number, count = 1): Sheet {
  const removable = Math.min(count, sheet.rowCount - at)
  return removable <= 0 ? sheet : shiftRows(sheet, at, -removable)
}

export function insertColumns(sheet: Sheet, at: number, count = 1): Sheet {
  return count <= 0 ? sheet : shiftColumns(sheet, at, count)
}

export function deleteColumns(sheet: Sheet, at: number, count = 1): Sheet {
  const removable = Math.min(count, sheet.columnCount - at)
  return removable <= 0 ? sheet : shiftColumns(sheet, at, -removable)
}

function shiftRows(sheet: Sheet, at: number, delta: number): Sheet {
  const cells: Record<string, Cell> = {}

  for (const [ref, cell] of Object.entries(sheet.cells)) {
    const position = positionOf(ref)
    if (position === null) continue

    if (position.row < at) {
      cells[ref] = cell
      continue
    }

    // A deleted row takes the cell with it.
    if (delta < 0 && position.row < at - delta) continue

    cells[cellRef(position.row + delta, position.column)] = cell
  }

  return {
    ...sheet,
    cells,
    rowHeights: shiftDimensions(sheet.rowHeights, at, delta),
    rowCount: Math.max(1, sheet.rowCount + delta),
    frozenRows: shiftFrozen(sheet.frozenRows, at, delta),
  }
}

function shiftColumns(sheet: Sheet, at: number, delta: number): Sheet {
  const cells: Record<string, Cell> = {}

  for (const [ref, cell] of Object.entries(sheet.cells)) {
    const position = positionOf(ref)
    if (position === null) continue

    if (position.column < at) {
      cells[ref] = cell
      continue
    }

    if (delta < 0 && position.column < at - delta) continue

    cells[cellRef(position.row, position.column + delta)] = cell
  }

  return {
    ...sheet,
    cells,
    columnWidths: shiftDimensions(sheet.columnWidths, at, delta),
    columnCount: Math.max(1, sheet.columnCount + delta),
    frozenColumns: shiftFrozen(sheet.frozenColumns, at, delta),
  }
}

/** Inserting inside the frozen band shifts it along. */
function shiftFrozen(frozen: number, at: number, delta: number): number {
  return at < frozen ? Math.max(at, frozen + delta) : frozen
}

function shiftDimensions(sizes: Record<number, number>, at: number, delta: number): Record<number, number> {
  const shifted: Record<number, number> = {}

  for (const [key, size] of Object.entries(sizes)) {
    const index = Number(key)
    if (index < at) {
      shifted[index] = size
      continue
    }
    if (delta < 0 && index < at - delta) continue
    shifted[index + delta] = size
  }

  return shifted
}

function positionOf(ref: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)([0-9]+)$/.exec(ref)
  if (match === null) return null

  let column = 0
  for (const letter of match[1]!) column = column * 26 + (letter.charCodeAt(0) - 64)

  return { row: Number.parseInt(match[2]!, 10) - 1, column: column - 1 }
}

function clean(style: Record<string, unknown>): CellStyle {
  const result: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(style)) {
    if (value !== undefined) result[key] = value
  }
  return result as CellStyle
}

function toCell(cell: Cell, style: CellStyle): Cell {
  const base: Cell = {}
  const withValue = cell.value === undefined ? base : { ...base, value: cell.value }
  const withFormula = cell.formula === undefined ? withValue : { ...withValue, formula: cell.formula }

  return Object.keys(style).length === 0 ? withFormula : { ...withFormula, style }
}
