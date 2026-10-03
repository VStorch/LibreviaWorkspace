import { FormulaError, isFormulaError } from '../errors.js'
import { compare, toBoolean, type Scalar } from '../values.js'
import { define, numberArg, rowsOf, single, type FunctionDefinition } from './kit.js'

/**
 * A busca "aproximada" do `PROCV` e do `CORRESP`: pressupõe o vetor crescente e
 * não confere, como no Excel, para o mesmo arquivo dar o mesmo resultado.
 */
function approximate(values: readonly Scalar[], target: Scalar): number {
  let found = -1
  for (let at = 0; at < values.length; at++) {
    const value = values[at] ?? null
    if (value === null) continue
    if (compare(value, target) <= 0) found = at
    else break
  }
  return found
}

function exact(values: readonly Scalar[], target: Scalar): number {
  return values.findIndex((value) => value !== null && compare(value, target) === 0)
}

export const LOOKUP: readonly FunctionDefinition[] = [
  /** O quarto argumento é VERDADEIRO por padrão, como no Excel, por pior que seja. */
  define(['PROCV', 'VLOOKUP'], 3, 4, (args) => {
    const target = single(args[0])
    if (isFormulaError(target)) return target
    const table = rowsOf(args[1])
    if (isFormulaError(table)) return table
    const column = numberArg(args[2])
    if (isFormulaError(column)) return column

    const wanted = Math.trunc(column)
    if (wanted < 1) return FormulaError.Value
    if (table.length === 0 || wanted > (table[0]?.length ?? 0)) return FormulaError.Ref

    const flag = args.length > 3 ? toBoolean(single(args[3])) : true
    if (isFormulaError(flag)) return flag

    const first = table.map((row) => row[0] ?? null)
    const at = flag ? approximate(first, target) : exact(first, target)
    if (at < 0) return FormulaError.NA

    return table[at]?.[wanted - 1] ?? null
  }),

  define(['PROCH', 'HLOOKUP'], 3, 4, (args) => {
    const target = single(args[0])
    if (isFormulaError(target)) return target
    const table = rowsOf(args[1])
    if (isFormulaError(table)) return table
    const row = numberArg(args[2])
    if (isFormulaError(row)) return row

    const wanted = Math.trunc(row)
    if (wanted < 1) return FormulaError.Value
    if (wanted > table.length) return FormulaError.Ref

    const flag = args.length > 3 ? toBoolean(single(args[3])) : true
    if (isFormulaError(flag)) return flag

    const first = table[0] ?? []
    const at = flag ? approximate(first, target) : exact(first, target)
    if (at < 0) return FormulaError.NA

    return table[wanted - 1]?.[at] ?? null
  }),

  /** Devolve a posição, e não o valor. */
  define(['CORRESP', 'MATCH'], 2, 3, (args) => {
    const target = single(args[0])
    if (isFormulaError(target)) return target
    const table = rowsOf(args[1])
    if (isFormulaError(table)) return table

    const kind = args.length > 2 ? numberArg(args[2]) : 1
    if (isFormulaError(kind)) return kind

    const values = table.flatMap((row) => [...row])

    if (kind === 0) {
      const at = exact(values, target)
      return at < 0 ? FormulaError.NA : at + 1
    }

    if (kind > 0) {
      const at = approximate(values, target)
      return at < 0 ? FormulaError.NA : at + 1
    }

    // Tipo negativo: vetor decrescente, primeiro valor maior ou igual.
    let found = -1
    for (let at = 0; at < values.length; at++) {
      const value = values[at] ?? null
      if (value === null) continue
      if (compare(value, target) >= 0) found = at
      else break
    }
    return found < 0 ? FormulaError.NA : found + 1
  }),

  define(['ÍNDICE', 'INDICE', 'INDEX'], 2, 3, (args) => {
    const table = rowsOf(args[0])
    if (isFormulaError(table)) return table
    const row = numberArg(args[1])
    if (isFormulaError(row)) return row
    const column = args.length > 2 ? numberArg(args[2]) : 1
    if (isFormulaError(column)) return column

    const wantedRow = Math.trunc(row)
    const wantedColumn = Math.trunc(column)
    if (wantedRow < 1 || wantedColumn < 1) return FormulaError.Value

    // Vetor de uma coluna com índice único: ÍNDICE(A1:A9;3) é a terceira célula.
    const line = table[wantedRow - 1]
    if (line === undefined || wantedColumn > line.length) return FormulaError.Ref

    return line[wantedColumn - 1] ?? null
  }),
]
