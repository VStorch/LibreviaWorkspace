/**
 * Gerado do **modelo**, e não do DOM como no documento: a grade só desenha as
 * células visíveis. As regras de formatação são as da tela (`formatCell`).
 */

import { formatCell } from './format.js'
import { cellRef, DEFAULT_COLUMN_WIDTH, type Cell, type Sheet } from './model.js'
import { translate, Language } from '@shared/i18n/index.js'

export interface PrintBounds {
  readonly rows: number
  readonly columns: number
}

/** Uma planilha nova tem mil linhas por vinte e seis colunas e nenhum dado. */
export function usedBounds(sheet: Sheet): PrintBounds {
  let rows = 0
  let columns = 0

  for (const reference of Object.keys(sheet.cells)) {
    const match = /^([A-Z]+)([0-9]+)$/.exec(reference)
    if (match === null) continue

    rows = Math.max(rows, Number(match[2]))
    columns = Math.max(columns, indexOfColumn(match[1]!) + 1)
  }

  return { rows, columns }
}

function indexOfColumn(letters: string): number {
  let index = 0
  for (const letter of letters) index = index * 26 + (letter.charCodeAt(0) - 64)
  return index - 1
}

/** As linhas congeladas viram `<thead>`, que o navegador repete no topo de cada página. */
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
 * O "ajustar à página" do Excel: em corpo 11, doze colunas numa A4 em retrato
 * partiriam números no meio. Quem quiser o texto grande usa paisagem.
 */
function fontSize(columns: number): number {
  if (columns <= 8) return 11
  if (columns <= 12) return 9
  if (columns <= 18) return 8
  return 7
}

/** Em proporção, e não em pixels, para a tabela caber na folha com as proporções da tela. */
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

/** Estilo direto, e não classes: cor e fundo são valores livres do arquivo. */
function inlineStyle(cell: Cell | undefined): string {
  const style = cell?.style
  if (style === undefined) return ''

  const parts: string[] = []
  if (style.bold === true) parts.push('font-weight:700')
  if (style.italic === true) parts.push('font-style:italic')
  if (style.underline === true) parts.push('text-decoration:underline')
  if (style.color !== undefined) parts.push(`color:${cssColor(style.color)}`)
  if (style.background !== undefined) parts.push(`background:${cssColor(style.background)}`)

  // Sem regra própria para número, como a grade: o papel não pode sair
  // diferente da tela.
  if (style.align !== undefined) parts.push(`text-align:${style.align}`)

  for (const side of style.borders ?? []) {
    parts.push(`border-${side}:1px solid #333`)
  }

  return parts.join(';')
}

/** O valor vem do arquivo: sem a trava, `red;background:url(...)` viraria outra declaração. */
function cssColor(value: string): string {
  return /^#[0-9a-fA-F]{6}$/.test(value) ? value : 'inherit'
}

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/** `break-inside: avoid`: a linha alta não é cortada pela quebra de página. */
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
