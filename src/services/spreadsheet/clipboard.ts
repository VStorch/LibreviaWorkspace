import { normalizeRange, writeText, type Range } from './edit.js'
import { formatCell } from './format.js'
import { getCell, type Sheet } from './model.js'

const COLUMN_SEPARATOR = '\t'
const ROW_SEPARATOR = '\n'
const QUOTE = '"'
const NEEDS_QUOTES = /[\t\n\r"]/u

/** As Excel and LibreOffice exchange cells: tabs between columns, a line per row. */
export function rangeAsText(sheet: Sheet, range: Range): string {
  const { fromRow, fromColumn, toRow, toColumn } = normalizeRange(range)
  const rows: string[] = []
  for (let row = fromRow; row <= toRow; row++) {
    const cells: string[] = []
    for (let column = fromColumn; column <= toColumn; column++) {
      cells.push(quoted(formatCell(getCell(sheet, row, column))))
    }
    rows.push(cells.join(COLUMN_SEPARATOR))
  }
  return rows.join(ROW_SEPARATOR)
}

function quoted(text: string): string {
  return NEEDS_QUOTES.test(text) ? QUOTE + text.replaceAll(QUOTE, QUOTE + QUOTE) + QUOTE : text
}

/** The rows and cells of copied text; a quoted cell may hold tabs and line breaks. */
export function parseClipboardText(text: string): string[][] {
  const rows: string[][] = [[]]
  let index = 0
  while (index <= text.length) {
    const [cell, next] = text[index] === QUOTE ? readQuoted(text, index + 1) : readPlain(text, index)
    rows.at(-1)!.push(cell)
    const separator = text[next]
    if (separator === '\r' && text[next + 1] === '\n') index = next + 2
    else index = next + 1
    if (separator !== COLUMN_SEPARATOR && index <= text.length) rows.push([])
  }
  const last = rows.at(-1)!
  if (rows.length > 1 && last.length === 1 && last[0] === '') rows.pop()
  return rows
}

function readPlain(text: string, start: number): [string, number] {
  let end = start
  while (end < text.length && !isSeparator(text[end]!)) end++
  return [text.slice(start, end), end]
}

function readQuoted(text: string, start: number): [string, number] {
  let cell = ''
  let index = start
  while (index < text.length) {
    if (text[index] !== QUOTE) cell += text[index++]
    else if (text[index + 1] === QUOTE) {
      cell += QUOTE
      index += 2
    } else return [cell, index + 1]
  }
  return [cell, index]
}

const isSeparator = (char: string): boolean => char === COLUMN_SEPARATOR || char === '\n' || char === '\r'

/** From the top left corner, as typed; whatever falls outside the sheet is dropped. */
export function pasteText(sheet: Sheet, row: number, column: number, rows: readonly string[][]): Sheet {
  let updated = sheet
  rows.forEach((cells, rowOffset) => {
    if (row + rowOffset >= sheet.rowCount) return
    cells.forEach((text, columnOffset) => {
      if (column + columnOffset >= sheet.columnCount) return
      updated = writeText(updated, row + rowOffset, column + columnOffset, text)
    })
  })
  return updated
}
