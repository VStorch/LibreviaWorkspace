/**
 * Árvore → valor.
 *
 * O avaliador não conhece a planilha: ele pede os valores ao contexto. É o que
 * permite testá-lo com um punhado de células de mentira, e é o que deixa o
 * recálculo decidir a ordem sem que a avaliação precise saber que existe ordem.
 *
 * Erro aqui é **valor**, não exceção — `#DIV/0!` se propaga pela conta como um
 * número se propagaria. Só erro de escrita da fórmula lança.
 */

import type { Node } from './ast.js'
import { FormulaError, isFormulaError } from './errors.js'
import type { CellRef } from './references.js'
import { compare, toBoolean, toNumber, toText, type Scalar } from './values.js'
import { findFunction } from './functions/index.js'

export interface EvalContext {
  /** Valor já calculado de uma célula. Célula vazia devolve `null`. */
  readonly valueAt: (ref: CellRef) => Scalar
  /**
   * O "agora" das funções de data.
   *
   * Injetado, e não lido do relógio: sem isso `HOJE()` tornaria todo teste
   * dependente do dia em que roda.
   */
  readonly now: () => Date
}

/** Um argumento de função: um valor, ou um retângulo de valores. */
export type Argument =
  | { readonly kind: 'value'; readonly value: Scalar }
  | { readonly kind: 'range'; readonly rows: readonly (readonly Scalar[])[] }

export function evaluate(node: Node, context: EvalContext): Scalar {
  switch (node.kind) {
    case 'number':
    case 'text':
    case 'boolean':
    case 'error':
      return node.value

    case 'reference':
      return context.valueAt(node.ref)

    // Um intervalo só faz sentido como argumento de função. Solto numa conta,
    // o Excel tenta uma interseção implícita que quase ninguém usa de propósito.
    case 'range':
      return FormulaError.Value

    case 'unary': {
      const value = toNumber(evaluate(node.operand, context))
      if (isFormulaError(value)) return value
      return node.operator === '-' ? -value : value
    }

    case 'percent': {
      const value = toNumber(evaluate(node.operand, context))
      return isFormulaError(value) ? value : value / 100
    }

    case 'binary':
      return binary(node.operator, node.left, node.right, context)

    case 'call':
      return call(node, context)
  }
}

function binary(operator: string, leftNode: Node, rightNode: Node, context: EvalContext): Scalar {
  const left = evaluate(leftNode, context)
  if (isFormulaError(left)) return left
  const right = evaluate(rightNode, context)
  if (isFormulaError(right)) return right

  switch (operator) {
    case '&': {
      const a = toText(left)
      if (isFormulaError(a)) return a
      const b = toText(right)
      return isFormulaError(b) ? b : a + b
    }

    case '=':
      return compare(left, right) === 0
    case '<>':
      return compare(left, right) !== 0
    case '<':
      return compare(left, right) < 0
    case '<=':
      return compare(left, right) <= 0
    case '>':
      return compare(left, right) > 0
    case '>=':
      return compare(left, right) >= 0

    default:
      return arithmetic(operator, left, right)
  }
}

function arithmetic(operator: string, left: Scalar, right: Scalar): Scalar {
  const a = toNumber(left)
  if (isFormulaError(a)) return a
  const b = toNumber(right)
  if (isFormulaError(b)) return b

  switch (operator) {
    case '+':
      return finite(a + b)
    case '-':
      return finite(a - b)
    case '*':
      return finite(a * b)
    case '/':
      // Divisão por zero é o erro que o usuário mais vê, e o que ele espera ver.
      return b === 0 ? FormulaError.Div0 : finite(a / b)
    case '^': {
      const power = a ** b
      return Number.isNaN(power) ? FormulaError.Num : finite(power)
    }
    default:
      return FormulaError.Value
  }
}

/** Estouro vira `#NÚM!`: `Infinity` numa célula não diz nada a ninguém. */
function finite(value: number): Scalar {
  return Number.isFinite(value) ? value : FormulaError.Num
}

/**
 * Três funções precisam receber os argumentos **sem avaliar**.
 *
 * `=SE(A1=0;"";1/A1)` não pode calcular `1/A1` quando A1 é zero: o ramo não
 * escolhido produziria `#DIV/0!` e a fórmula inteira erraria. O mesmo vale para
 * `SEERRO`, cujo propósito é justamente não deixar o erro passar adiante.
 */
const LAZY = new Set(['SE', 'IF', 'SEERRO', 'IFERROR', 'SENÃODISP', 'SEND', 'IFNA'])

function call(node: { name: string; args: readonly Node[] }, context: EvalContext): Scalar {
  if (LAZY.has(node.name)) return lazyCall(node.name, node.args, context)

  const definition = findFunction(node.name)
  if (definition === undefined) return FormulaError.Name
  if (node.args.length < definition.minArgs || node.args.length > definition.maxArgs) {
    return FormulaError.Value
  }

  const args: Argument[] = []
  for (const arg of node.args) {
    const value = argumentOf(arg, context)
    // Erro em argumento contamina a chamada inteira — menos nas preguiçosas
    // acima e nas que existem justamente para examinar o erro.
    if (!definition.acceptsErrors && value.kind === 'value' && isFormulaError(value.value)) {
      return value.value
    }
    args.push(value)
  }

  return definition.call(args, context)
}

function lazyCall(name: string, args: readonly Node[], context: EvalContext): Scalar {
  if (name === 'SE' || name === 'IF') {
    if (args.length < 2 || args.length > 3) return FormulaError.Value

    const condition = toBoolean(evaluate(args[0]!, context))
    if (isFormulaError(condition)) return condition

    const branch = condition ? args[1] : args[2]
    // `=SE(A1>0;1)` com A1 negativo devolve FALSO no Excel, e não vazio.
    return branch === undefined ? false : evaluate(branch, context)
  }

  if (args.length !== 2) return FormulaError.Value
  const value = evaluate(args[0]!, context)

  const intercepts =
    name === 'SEERRO' || name === 'IFERROR' ? isFormulaError(value) : value === FormulaError.NA
  return intercepts ? evaluate(args[1]!, context) : value
}

function argumentOf(node: Node, context: EvalContext): Argument {
  if (node.kind !== 'range') return { kind: 'value', value: evaluate(node, context) }

  const rows: Scalar[][] = []
  for (let row = node.from.row; row <= node.to.row; row++) {
    const line: Scalar[] = []
    for (let column = node.from.column; column <= node.to.column; column++) {
      const ref: CellRef = {
        row,
        column,
        rowAbsolute: false,
        columnAbsolute: false,
        ...(node.from.sheet === undefined ? {} : { sheet: node.from.sheet }),
      }
      line.push(context.valueAt(ref))
    }
    rows.push(line)
  }

  return { kind: 'range', rows }
}
