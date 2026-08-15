/**
 * Modelo canônico da planilha.
 *
 * Espelha `document/model.ts`: dado puro e serializável, sem nada de React nem
 * de Electron. A Fase 7 vai mapear XLSX para cá, e o motor de fórmulas da Fase
 * 6 vai ler daqui — por isso o formato precisa ser estável antes da interface.
 */

/** Como o valor é exibido. O valor cru fica sempre separado da aparência. */
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

export interface CellStyle {
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly color?: string
  readonly background?: string
  readonly align?: HorizontalAlign
  readonly format?: CellFormat
  /** Casas decimais para número, moeda e percentual. */
  readonly decimals?: number
  readonly borders?: readonly BorderSide[]
}

/** O dado cru de uma célula, sem aparência. */
export type CellValue = string | number | boolean

/**
 * O conteúdo de uma célula.
 *
 * `value` é o dado cru — número, texto ou booleano. `formula` guarda a fórmula
 * digitada, quando há; o valor calculado continua em `value`. Separar os dois
 * é o que permite reabrir a planilha sem recalcular tudo, e é como o XLSX
 * também guarda.
 */
export interface Cell {
  readonly value?: CellValue
  readonly formula?: string
  readonly style?: CellStyle
}

/**
 * Células indexadas por referência A1.
 *
 * Mapa esparso, e não matriz: uma planilha de 10 mil linhas com trinta células
 * preenchidas ocupa trinta entradas. Matriz densa gastaria memória proporcional
 * à área, não ao conteúdo — e a área é o que o usuário rola, não o que ele
 * digita.
 */
export type CellMap = Record<string, Cell>

export interface Sheet {
  readonly name: string
  readonly cells: CellMap
  /** Larguras em pixels, por índice de coluna. Ausente = padrão. */
  readonly columnWidths: Record<number, number>
  readonly rowHeights: Record<number, number>
  /** Quantas linhas e colunas ficam presas ao rolar. */
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

// --- referências A1 --------------------------------------------------------

/**
 * Índice de coluna → letra: 0 → A, 25 → Z, 26 → AA.
 *
 * A base 26 do Excel não tem zero: depois de Z vem AA, não BA. Por isso o
 * decremento antes de cada divisão.
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

/** Referência A1 de uma posição base zero. */
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
 * Devolve a planilha com uma célula alterada.
 *
 * Célula que fica sem valor, sem fórmula e sem estilo é **removida** do mapa em
 * vez de guardada vazia: senão apagar o conteúdo de uma área deixaria milhares
 * de entradas inúteis, e o arquivo cresceria a cada limpeza.
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
