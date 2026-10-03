/**
 * Como no Excel, dentro de um intervalo texto e booleano são **ignorados** por
 * `SOMA`, e passados direto são convertidos: a coluna com cabeçalho de texto soma.
 */

import type { Argument, EvalContext } from '../evaluate.js'
import { FormulaError, isFormulaError } from '../errors.js'
import { toNumber, toText, type Scalar } from '../values.js'

export interface FunctionDefinition {
  /** O português primeiro, o inglês depois. */
  readonly names: readonly string[]
  readonly minArgs: number
  readonly maxArgs: number
  /** Só para quem examina o erro: sem isto `ÉERROS(A1)` nunca seria chamada. */
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

/** Vazio é sempre ignorado: `MÉDIA(A1;A2)` com A1 vazia divide por um. */
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

/** Todos os valores, sem conversão — para contar e para procurar. */
export function valuesIn(args: readonly Argument[]): Scalar[] {
  const values: Scalar[] = []
  for (const arg of args) {
    if (arg.kind === 'range') for (const row of arg.rows) values.push(...row)
    else values.push(arg.value)
  }
  return values
}

/** Intervalo aqui é erro de escrita: `=ARRED(A1:B2;2)` não quer dizer nada. */
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

/** O retângulo de um argumento de intervalo, ou o valor solto como 1×1. */
export function rowsOf(arg: Argument | undefined): readonly (readonly Scalar[])[] | FormulaError {
  if (arg === undefined) return FormulaError.Value
  return arg.kind === 'range' ? arg.rows : [[arg.value]]
}
