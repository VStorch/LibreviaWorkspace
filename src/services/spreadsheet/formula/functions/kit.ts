/**
 * Ferramentas comuns às funções de fórmula.
 *
 * As regras de coerção não são as óbvias: dentro de um intervalo, texto e
 * booleano são **ignorados** por `SOMA`; passados direto como argumento, são
 * convertidos. Parece inconsistência, mas é o que o Excel faz — e é o que faz
 * sentido, porque uma coluna com um cabeçalho de texto ainda deve somar.
 */

import type { Argument, EvalContext } from '../evaluate.js'
import { FormulaError, isFormulaError } from '../errors.js'
import { toNumber, toText, type Scalar } from '../values.js'

export interface FunctionDefinition {
  /** Nomes aceitos: o português primeiro, o inglês depois. */
  readonly names: readonly string[]
  readonly minArgs: number
  readonly maxArgs: number
  /**
   * Recebe argumentos com erro em vez de propagá-los.
   *
   * Vale só para quem existe para **examinar** o erro: sem isso `ÉERROS(A1)`
   * nunca seria chamada, porque o erro de A1 viraria o resultado antes.
   */
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

/**
 * Os números de um conjunto de argumentos, para as funções de agregação.
 *
 * Vazio é sempre ignorado, inclusive quando vem como referência solta: é o que
 * faz `MÉDIA(A1;A2)` com A1 vazia dividir por um, e não por dois.
 */
export function numbersIn(args: readonly Argument[]): number[] | FormulaError {
  const numbers: number[] = []

  for (const arg of args) {
    if (arg.kind === 'range') {
      for (const row of arg.rows) {
        for (const cell of row) {
          if (isFormulaError(cell)) return cell
          if (typeof cell === 'number') numbers.push(cell)
        }
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

/**
 * Um argumento que deveria ser um valor único.
 *
 * Intervalo aqui é erro de escrita: `=ARRED(A1:B2;2)` não quer dizer nada, e
 * devolver o primeiro da lista esconderia o engano.
 */
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
