import type { Range } from '@services/spreadsheet/edit.js'
import { formatCell } from '@services/spreadsheet/format.js'
import { cellRef, columnName, parseRef, type Cell, type Sheet } from '@services/spreadsheet/model.js'

export interface CellPosition {
  readonly row: number
  readonly column: number
}

interface UsedCell extends CellPosition {
  readonly cell: Cell
}

/** A long text would take the room of many cells; the start says what the cell holds. */
const MAX_CELL_TEXT = 200
const LINE_BREAKS = /[\r\n]+/g
const NEAREST_ONLY = 'Only the cells nearest the active one are shown.'
const ITEM_SEPARATOR = ' | '

function isUsed(cell: Cell): boolean {
  return cell.formula !== undefined || (cell.value !== undefined && cell.value !== '')
}

function usedCells(sheet: Sheet): UsedCell[] {
  const used: UsedCell[] = []
  for (const [ref, cell] of Object.entries(sheet.cells)) {
    const position = parseRef(ref)
    if (position !== null && isUsed(cell)) used.push({ ...position, cell })
  }
  return used
}

/** One line per cell: a line break inside a cell would read as another cell. */
function shown(text: string): string {
  return text.replace(LINE_BREAKS, ' ').slice(0, MAX_CELL_TEXT)
}

function cellText(cell: Cell): string {
  const value = shown(formatCell(cell))
  return cell.formula === undefined ? value : `${shown(cell.formula)} → ${value}`
}

const item = (used: UsedCell): string => `${cellRef(used.row, used.column)}: ${cellText(used.cell)}`

const byPosition = (a: CellPosition, b: CellPosition): number => a.row - b.row || a.column - b.column

/** The first row with anything in it: where a table keeps its titles. */
function headerRowOf(used: readonly UsedCell[]): number | null {
  return used.length === 0 ? null : Math.min(...used.map((cell) => cell.row))
}

function headerCells(used: readonly UsedCell[], headerRow: number | null): UsedCell[] {
  return used.filter((cell) => cell.row === headerRow).sort(byPosition)
}

const distance = (cell: CellPosition, active: CellPosition): number =>
  Math.abs(cell.row - active.row) + Math.abs(cell.column - active.column)

/** Whole lines while they fit, in order; says so when some were left out. */
export function fitLines(lines: readonly string[], budget: number): string {
  const kept: string[] = []
  let used = 0
  for (const line of lines) {
    if (used + line.length + 1 > budget) {
      const withNote = [...kept, NEAREST_ONLY].join('\n')
      return withNote.length <= budget ? withNote : kept.join('\n')
    }
    kept.push(line)
    used += line.length + 1
  }
  return kept.join('\n')
}

/**
 * For explaining and writing a formula: the sheet's name, its headers and the used cells, nearest
 * to the active one first, as many as the budget holds. Other sheets never go.
 */
export function describeSheetAround(sheet: Sheet, active: CellPosition, budget: number): string {
  const used = usedCells(sheet)
  const headerRow = headerRowOf(used)
  const headers = headerCells(used, headerRow)
  const nearby = used
    .filter((cell) => cell.row !== headerRow)
    .sort((a, b) => distance(a, active) - distance(b, active) || byPosition(a, b))
  const lines = [
    `Sheet: ${sheet.name}`,
    `Active cell: ${cellRef(active.row, active.column)}`,
    ...(headers.length === 0 ? [] : ['Headers:', ...headers.map(item)]),
    ...(nearby.length === 0 ? [] : ['Cells:', ...nearby.map(item)]),
  ]
  return fitLines(lines, budget)
}

function columnLabel(column: number, headers: readonly UsedCell[]): string {
  const header = headers.find((cell) => cell.column === column)
  const letter = columnName(column)
  return header === undefined ? letter : `${letter} (${shown(formatCell(header.cell))})`
}

function groupByRow(cells: readonly UsedCell[]): Map<number, UsedCell[]> {
  const rows = new Map<number, UsedCell[]>()
  for (const cell of cells) {
    const row = rows.get(cell.row)
    if (row === undefined) rows.set(cell.row, [cell])
    else row.push(cell)
  }
  return rows
}

function rowLine(row: number, cells: readonly UsedCell[]): string {
  const values = cells.map((cell) => `${cellRef(cell.row, cell.column)}: ${shown(formatCell(cell.cell))}`)
  return `Row ${row + 1}: ${values.length === 0 ? '(empty)' : values.join(ITEM_SEPARATOR)}`
}

/**
 * For filling a column: each target row with the values of the other used columns, and their
 * headers when the first row with content is above the target. Never cut: a missing row would
 * misplace every value after it.
 */
export function describeRowsToFill(sheet: Sheet, target: Range): string {
  const used = usedCells(sheet)
  const firstRow = headerRowOf(used)
  const headerRow = firstRow !== null && firstRow < target.fromRow ? firstRow : null
  const headers = headerCells(used, headerRow)
  const others = used.filter((cell) => cell.column !== target.fromColumn)
  const columns = [...new Set(others.map((cell) => cell.column))].sort((a, b) => a - b)
  const byRow = groupByRow(others)
  const rows: string[] = []
  for (let row = target.fromRow; row <= target.toRow; row++) {
    rows.push(rowLine(row, [...(byRow.get(row) ?? [])].sort(byPosition)))
  }
  return [
    `Sheet: ${sheet.name}`,
    `Target column: ${columnLabel(target.fromColumn, headers)}`,
    ...(columns.length === 0
      ? []
      : [`Columns: ${columns.map((c) => columnLabel(c, headers)).join(ITEM_SEPARATOR)}`]),
    `Rows (${target.toRow - target.fromRow + 1}):`,
    ...rows,
  ].join('\n')
}
