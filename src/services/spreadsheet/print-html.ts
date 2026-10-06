import { escapeHtml } from '@services/html.js'
/**
 * Built from the **model**, not from the DOM as for documents: the grid only draws visible cells.
 * Formatting rules are the screen's (`formatCell`).
 */

import { formatCell } from './format.js'
import { cellRef, columnIndex, DEFAULT_COLUMN_WIDTH, type Cell, type Sheet } from './model.js'
import { translate, Language } from '@shared/i18n/index.js'

export interface PrintBounds {
  readonly rows: number
  readonly columns: number
}

/** A new sheet has a thousand rows by twenty-six columns and no data. */
export function usedBounds(sheet: Sheet): PrintBounds {
  let rows = 0
  let columns = 0

  for (const reference of Object.keys(sheet.cells)) {
    const match = /^([A-Z]+)([0-9]+)$/.exec(reference)
    if (match === null) continue

    rows = Math.max(rows, Number(match[2]))
    columns = Math.max(columns, columnIndex(match[1]!) + 1)
  }

  return { rows, columns }
}

/** Frozen rows become `<thead>`, which the browser repeats at the top of each page. */
export function buildSheetHtml(sheet: Sheet, language: Language = Language.Portuguese): string {
  const bounds = usedBounds(sheet)
  if (bounds.rows === 0 || bounds.columns === 0) {
    return `<p class="sheet-print__empty">${translate(language, 'spreadsheet.print.emptyTab')}</p>`
  }

  const widths = columnWidths(sheet, bounds)

  const frozen = Math.min(sheet.frozenRows, bounds.rows)
  const head = frozen > 0 ? `<thead>${rowsHtml(sheet, bounds, 0, frozen)}</thead>` : ''
  const body = `<tbody>${rowsHtml(sheet, bounds, frozen, bounds.rows)}</tbody>`

  return `<table class="sheet-print" style="font-size:${fontSize(bounds.columns)}pt"><colgroup>${widths}</colgroup>${head}${body}</table>`
}

/**
 * Excel's "fit to page": at 11 pt, twelve columns on portrait A4 would break numbers in the middle.
 * Whoever wants large text uses landscape.
 */
function fontSize(columns: number): number {
  if (columns <= 8) return 11
  if (columns <= 12) return 9
  if (columns <= 18) return 8
  return 7
}

/** Proportional, not pixels, so the table fits the sheet with the screen's proportions. */
function columnWidths(sheet: Sheet, bounds: PrintBounds): string {
  const pixels = Array.from(
    { length: bounds.columns },
    (_, column) => sheet.columnWidths[column] ?? DEFAULT_COLUMN_WIDTH,
  )

  const total = pixels.reduce((sum, width) => sum + width, 0)
  return pixels.map((width) => `<col style="width:${((width / total) * 100).toFixed(3)}%">`).join('')
}

function rowsHtml(sheet: Sheet, bounds: PrintBounds, from: number, to: number): string {
  const rows: string[] = []

  for (let row = from; row < to; row++) {
    const cells: string[] = []
    for (let column = 0; column < bounds.columns; column++) {
      cells.push(cellHtml(sheet.cells[cellRef(row, column)]))
    }

    const height = sheet.rowHeights[row]
    const style = height === undefined ? '' : ` style="height:${Math.round(height)}px"`
    rows.push(`<tr${style}>${cells.join('')}</tr>`)
  }

  return rows.join('')
}

function cellHtml(cell: Cell | undefined): string {
  const text = escapeHtml(formatCell(cell))
  const style = inlineStyle(cell)
  return style === '' ? `<td>${text}</td>` : `<td style="${style}">${text}</td>`
}

/** Inline style, not classes: color and background are free values from the file. */
function inlineStyle(cell: Cell | undefined): string {
  const style = cell?.style
  if (style === undefined) return ''

  const parts: string[] = []
  if (style.bold === true) parts.push('font-weight:700')
  if (style.italic === true) parts.push('font-style:italic')
  if (style.underline === true) parts.push('text-decoration:underline')
  if (style.color !== undefined) parts.push(`color:${cssColor(style.color)}`)
  if (style.background !== undefined) parts.push(`background:${cssColor(style.background)}`)

  // No rule of its own for numbers, like the grid: paper must not differ from the screen.
  if (style.align !== undefined) parts.push(`text-align:${style.align}`)

  for (const side of style.borders ?? []) {
    parts.push(`border-${side}:1px solid #333`)
  }

  return parts.join(';')
}

/**
 * The value comes from the file: without the guard, `red;background:url(...)` would become another
 * declaration.
 */
function cssColor(value: string): string {
  return /^#[0-9a-fA-F]{6}$/.test(value) ? value : 'inherit'
}

/** `break-inside: avoid`: a tall row is not cut by the page break. */
export const SHEET_PRINT_CSS = `
.sheet-print {
  border-collapse: collapse;
  table-layout: fixed;
  width: 100%;
  font-family: Calibri, Carlito, system-ui, sans-serif;
  /* O corpo real vem no atributo da tabela: depende da largura da planilha. */
  font-size: 11pt;
}
/* O texto quebra em vez de ser cortado, porque no papel não dá para alargar a
   coluna; mas só entre palavras, porque "200" partido em "20" e "0" engana. */
.sheet-print td {
  border: 1px solid #d0d0d0;
  padding: 2px 5px;
  vertical-align: bottom;
  word-break: normal;
  overflow-wrap: break-word;
}
.sheet-print tr { break-inside: avoid; }
.sheet-print thead td { font-weight: 700; background: #f2f2f2; }
.sheet-print__empty { font-family: system-ui, sans-serif; color: #666; }
`
