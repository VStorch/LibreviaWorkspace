/**
 * While typing the user deserves the sentence, not `#VALOR!`: `=ABS(1;2)` has too many arguments.
 * Excel also refuses the formula with a warning.
 */

import { translate, Language } from '@shared/i18n/index.js'
import { walk, type Node } from './ast.js'
import { ParseError } from './errors.js'
import { findFunction, isKnownFunction } from './functions/index.js'
import { parseFormula } from './parse.js'

export interface FormulaProblem {
  readonly message: string
  /** Position in the formula, when the problem is in the writing. */
  readonly position?: number
}

export function checkFormula(
  formula: string,
  language: Language = Language.Portuguese,
): FormulaProblem | null {
  let node: Node
  try {
    node = parseFormula(formula)
  } catch (cause) {
    if (cause instanceof ParseError) return { message: cause.describe(language), position: cause.position }
    throw cause
  }

  for (const child of walk(node)) {
    if (child.kind !== 'call') continue

    if (!isKnownFunction(child.name)) {
      return { message: translate(language, 'spreadsheet.formula.unknownFunction', { name: child.name }) }
    }

    const problem = arityOf(child.name, child.args.length, language)
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

function arityOf(name: string, given: number, language: Language): string | null {
  const upper = name.toUpperCase()
  const limits = LAZY_ARITY[upper] ?? findFunction(upper)
  if (limits === undefined) return null

  const min = 'min' in limits ? limits.min : limits.minArgs
  const max = 'max' in limits ? limits.max : limits.maxArgs

  const count = (value: number) => translate(language, 'spreadsheet.formula.arguments', { count: value })
  if (given < min) {
    return translate(language, 'spreadsheet.formula.tooFewArguments', {
      name,
      expected: count(min),
      given: given === 0 ? translate(language, 'spreadsheet.formula.noArguments') : count(given),
    })
  }
  if (given > max) {
    return translate(language, 'spreadsheet.formula.tooManyArguments', { name, expected: count(max), given })
  }

  return null
}
