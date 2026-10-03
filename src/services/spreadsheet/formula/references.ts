/** O `$` decide o que acontece ao copiar: `A1` acompanha, `` não sai, `A` e `` prendem um eixo. */

import { columnIndex, columnName } from '../model.js'
import { ParseError } from './errors.js'

export interface CellRef {
  readonly sheet?: string | undefined
  readonly row: number
  readonly column: number
  readonly rowAbsolute: boolean
  readonly columnAbsolute: boolean
}

/** `Planilha1!$B$3` → referência. Devolve `null` se não for uma. */
export function parseReference(text: string): CellRef | null {
  const { sheet, rest } = splitSheet(text)
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(rest)
  if (match === null) return null

  const column = columnIndex(match[2]!)
  const row = Number.parseInt(match[4]!, 10) - 1
  if (column < 0 || row < 0) return null

  const ref: CellRef = {
    row,
    column,
    columnAbsolute: match[1] === '$',
    rowAbsolute: match[3] === '$',
  }
  return sheet === undefined ? ref : { ...ref, sheet }
}

export function formatReference(ref: CellRef): string {
  const body = `${ref.columnAbsolute ? '$' : ''}${columnName(ref.column)}${ref.rowAbsolute ? '$' : ''}${ref.row + 1}`
  return ref.sheet === undefined ? body : `${quoteSheet(ref.sheet)}!${body}`
}

/** Sem apóstrofos, `Vendas 2026!A1` se partiria no espaço. */
export function quoteSheet(name: string): string {
  return /^[\p{L}\p{N}_]+$/u.test(name) ? name : `'${name.replaceAll("'", "''")}'`
}

function splitSheet(text: string): { sheet?: string; rest: string } {
  if (text.startsWith("'")) {
    const end = text.indexOf("'!", 1)
    if (end < 0) throw new ParseError('Faltou fechar o apóstrofo do nome da planilha.', 0)
    return { sheet: text.slice(1, end).replaceAll("''", "'"), rest: text.slice(end + 2) }
  }

  const bang = text.indexOf('!')
  return bang < 0 ? { rest: text } : { sheet: text.slice(0, bang), rest: text.slice(bang + 1) }
}
