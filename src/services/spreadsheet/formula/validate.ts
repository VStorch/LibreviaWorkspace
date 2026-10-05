/**
 * While typing the user deserves the sentence, not `#VALOR!`: `=ABS(1;2)` has too many arguments.
 * Excel also refuses the formula with a warning.
 */

import { walk, type Node } from './ast.js'
import { ParseError } from './errors.js'
import { findFunction, isKnownFunction } from './functions/index.js'
import { parseFormula } from './parse.js'

export interface FormulaProblem {
  readonly message: string
  /** Position in the formula, when the problem is in the writing. */
  readonly position?: number
}

export function checkFormula(formula: string): FormulaProblem | null {
  let node: Node
  try {
    node = parseFormula(formula)
  } catch (cause) {
    if (cause instanceof ParseError) return { message: cause.message, position: cause.position }
    throw cause
  }

  for (const child of walk(node)) {
    if (child.kind !== 'call') continue

    if (!isKnownFunction(child.name)) {
      return { message: `Não conheço a função ${child.name}.` }
    }

    const problem = arityOf(child.name, child.args.length)
    if (problem !== null) return { message: problem }
  }

  return null
}

/** Lazy functions are not in the catalog. */
const LAZY_ARITY: Record<string, { min: number; max: number }> = {
  SE: { min: 2, max: 3 },
  IF: { min: 2, max: 3 },
  SEERRO: { min: 2, max: 2 },
  IFERROR: { min: 2, max: 2 },
  SENÃODISP: { min: 2, max: 2 },
  SEND: { min: 2, max: 2 },
  IFNA: { min: 2, max: 2 },
}

function arityOf(name: string, given: number): string | null {
  const upper = name.toUpperCase()
  const limits = LAZY_ARITY[upper] ?? findFunction(upper)
  if (limits === undefined) return null

  const min = 'min' in limits ? limits.min : limits.minArgs
  const max = 'max' in limits ? limits.max : limits.maxArgs

  if (given < min) {
    return `${name} precisa de ${count(min)}, e recebeu ${given === 0 ? 'nenhum' : count(given)}.`
  }
  if (given > max) {
    return `${name} aceita no máximo ${count(max)}, e recebeu ${given}.`
  }

  return null
}

function count(value: number): string {
  return value === 1 ? '1 argumento' : `${value} argumentos`
}
