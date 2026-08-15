/**
 * Referências dentro de uma fórmula.
 *
 * O `$` não é enfeite: ele decide o que acontece ao copiar a fórmula para outra
 * célula. `A1` acompanha o deslocamento, `$A$1` não sai do lugar, e `A$1` e
 * `$A1` prendem só um eixo. Guardar essa informação no modelo é o que permite
 * copiar e preencher sem reescrever nada.
 */

import { columnIndex, columnName } from '../model.js'
import { ParseError } from './errors.js'

export interface CellRef {
  /** Nome da planilha, quando a referência atravessa abas. */
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

/**
 * Nome de planilha que precisa de apóstrofos.
 *
 * Sem eles, `Vendas 2026!A1` se partiria no espaço e a fórmula deixaria de ser
 * lida de volta.
 */
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
