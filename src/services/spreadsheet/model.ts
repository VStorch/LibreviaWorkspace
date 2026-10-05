/** The raw value stays apart from the appearance. */
export const CellFormat = {
  General: 'general',
  Text: 'text',
  Number: 'number',
  Currency: 'currency',
  Percent: 'percent',
  Date: 'date',
} as const
export type CellFormat = (typeof CellFormat)[keyof typeof CellFormat]

export const HorizontalAlign = {
  Left: 'left',
  Center: 'center',
  Right: 'right',
} as const
export type HorizontalAlign = (typeof HorizontalAlign)[keyof typeof HorizontalAlign]

export const BorderSide = {
  Top: 'top',
  Right: 'right',
  Bottom: 'bottom',
  Left: 'left',
} as const
export type BorderSide = (typeof BorderSide)[keyof typeof BorderSide]

/**
 * Explicit `| undefined` to be assignable to what zod infers under `exactOptionalPropertyTypes`.
 */
export interface CellStyle {
  readonly bold?: boolean | undefined
  readonly italic?: boolean | undefined
  readonly underline?: boolean | undefined
  readonly color?: string | undefined
  readonly background?: string | undefined
  readonly align?: HorizontalAlign | undefined
  readonly format?: CellFormat | undefined
  /** For number, currency and percent. */
  readonly decimals?: number | undefined
  readonly borders?: readonly BorderSide[] | undefined
}

export type CellValue = string | number | boolean

/**
 * `value` holds the computed value and `formula` the formula: reopening does not require
 * recalculating, as in XLSX.
 */
export interface Cell {
  readonly value?: CellValue | undefined
  readonly formula?: string | undefined
  readonly style?: CellStyle | undefined
}

/** Sparse, not a matrix: ten thousand rows with thirty filled cells take thirty entries. */
export type CellMap = Record<string, Cell>

export interface Sheet {
  readonly name: string
  readonly cells: CellMap
  /** Pixels, by column index. Absent means default. */
  readonly columnWidths: Record<number, number>
  readonly rowHeights: Record<number, number>
  /** Rows and columns pinned while scrolling. */
  readonly frozenRows: number
  readonly frozenColumns: number
  readonly rowCount: number
  readonly columnCount: number
}

export interface WorkbookModel {
  readonly sheets: Sheet[]
  readonly activeSheet: number
}

export const DEFAULT_ROW_COUNT = 1000
export const DEFAULT_COLUMN_COUNT = 26
export const DEFAULT_COLUMN_WIDTH = 96
export const DEFAULT_ROW_HEIGHT = 24

export function createSheet(name: string): Sheet {
  return {
    name,
    cells: {},
    columnWidths: {},
    rowHeights: {},
    frozenRows: 0,
    frozenColumns: 0,
    rowCount: DEFAULT_ROW_COUNT,
    columnCount: DEFAULT_COLUMN_COUNT,
  }
}

export function createEmptyWorkbook(): WorkbookModel {
  return { sheets: [createSheet('Planilha1')], activeSheet: 0 }
}

/**
 * 0 → A, 25 → Z, 26 → AA. Excel's base 26 has no zero, hence the decrement before each division.
 */
export function columnName(index: number): string {
  if (!Number.isInteger(index) || index < 0) return ''

  let name = ''
  let remaining = index
  while (remaining >= 0) {
    name = String.fromCharCode(65 + (remaining % 26)) + name
    remaining = Math.floor(remaining / 26) - 1
  }
  return name
}

export function columnIndex(name: string): number {
  const letters = name.toUpperCase()
  if (!/^[A-Z]+$/.test(letters)) return -1

  let index = 0
  for (const letter of letters) {
    index = index * 26 + (letter.charCodeAt(0) - 64)
  }
  return index - 1
}

export function cellRef(row: number, column: number): string {
  return `${columnName(column)}${row + 1}`
}

export function parseRef(ref: string): { row: number; column: number } | null {
  const match = /^([A-Za-z]+)([0-9]+)$/.exec(ref.trim())
  if (match === null) return null

  const column = columnIndex(match[1]!)
  const row = Number.parseInt(match[2]!, 10) - 1
  if (column < 0 || row < 0) return null

  return { row, column }
}

export function getCell(sheet: Sheet, row: number, column: number): Cell | undefined {
  return sheet.cells[cellRef(row, column)]
}

/**
 * A cell without value, formula or style leaves the map, so the file does not grow with every
 * clear.
 */
export function setCell(sheet: Sheet, row: number, column: number, cell: Cell): Sheet {
  const ref = cellRef(row, column)
  const cells = { ...sheet.cells }

  if (isBlank(cell)) delete cells[ref]
  else cells[ref] = cell

  return { ...sheet, cells }
}

export function isBlank(cell: Cell): boolean {
  const emptyValue = cell.value === undefined || cell.value === ''
  const emptyStyle = cell.style === undefined || Object.keys(cell.style).length === 0
  return emptyValue && cell.formula === undefined && emptyStyle
}
