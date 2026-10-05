/**
 * As in Excel, inside a range text and booleans are **ignored** by `SOMA`, and passed directly they
 * are converted: a column with a text header sums.
 */

import type { Argument, EvalContext } from '../evaluate.js'
import { FormulaError, isFormulaError } from '../errors.js'
import { toNumber, toText, type Scalar } from '../values.js'

export interface FunctionDefinition {
  /** Portuguese first, English after. */
  readonly names: readonly string[]
  readonly minArgs: number
  readonly maxArgs: number
  /** Only for functions that inspect errors: otherwise `ÉERROS(A1)` would never be called. */
  readonly acceptsErrors: boolean
  readonly call: (args: readonly Argument[], context: EvalContext) => Scalar
}

export function define(
  names: readonly string[],
  minArgs: number,
  maxArgs: number,
  call: FunctionDefinition['call'],
  acceptsErrors = false,
): FunctionDefinition {
  return { names, minArgs, maxArgs, call, acceptsErrors }
}

export const VARIADIC = Number.MAX_SAFE_INTEGER

/** Empty is always ignored: `MÉDIA(A1;A2)` with A1 empty divides by one. */
export function numbersIn(args: readonly Argument[]): number[] | FormulaError {
  const numbers: number[] = []

  for (const arg of args) {
    if (arg.kind === 'range') {
      for (const cell of arg.rows.flat()) {
        if (isFormulaError(cell)) return cell
        if (typeof cell === 'number') numbers.push(cell)
      }
      continue
    }

    if (arg.value === null) continue
    const number = toNumber(arg.value)
    if (isFormulaError(number)) return number
    numbers.push(number)
  }

  return numbers
}

/** All values, unconverted: for counting and lookups. */
export function valuesIn(args: readonly Argument[]): Scalar[] {
  const values: Scalar[] = []
  for (const arg of args) {
    if (arg.kind === 'range') for (const row of arg.rows) values.push(...row)
    else values.push(arg.value)
  }
  return values
}

/** A range here is a malformed formula: `=ARRED(A1:B2;2)` means nothing. */
export function single(arg: Argument | undefined): Scalar {
  if (arg === undefined) return null
  return arg.kind === 'value' ? arg.value : FormulaError.Value
}

export function numberArg(arg: Argument | undefined): number | FormulaError {
  return toNumber(single(arg))
}

export function textArg(arg: Argument | undefined): string | FormulaError {
  return toText(single(arg))
}

/** A range argument's rectangle, or a loose value as 1×1. */
export function rowsOf(arg: Argument | undefined): readonly (readonly Scalar[])[] | FormulaError {
  if (arg === undefined) return FormulaError.Value
  return arg.kind === 'range' ? arg.rows : [[arg.value]]
}
