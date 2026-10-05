/** `SE` and `SEERRO` receive their arguments unevaluated and live in the evaluator. */

import type { Argument } from '../evaluate.js'
import { FormulaError, isFormulaError } from '../errors.js'
import { toBoolean, type Scalar } from '../values.js'
import { VARIADIC, define, single, valuesIn, type FunctionDefinition } from './kit.js'

/** As in Excel, they evaluate every argument: `=E(FALSO;1/0)` is `#DIV/0!`. */
function fold(
  args: readonly Argument[],
  combine: (a: boolean, b: boolean) => boolean,
  seed: boolean,
): Scalar {
  let result = seed
  let seen = false

  for (const value of valuesIn(args)) {
    // Empty inside a range is ignored, as in Excel.
    if (value === null) continue

    const flag = toBoolean(value)
    if (isFormulaError(flag)) return flag
    result = combine(result, flag)
    seen = true
  }

  return seen ? result : FormulaError.Value
}

export const LOGICAL: readonly FunctionDefinition[] = [
  define(['E', 'AND'], 1, VARIADIC, (args) => fold(args, (a, b) => a && b, true)),
  define(['OU', 'OR'], 1, VARIADIC, (args) => fold(args, (a, b) => a || b, false)),

  define(['NÃO', 'NAO', 'NOT'], 1, 1, (args) => {
    const flag = toBoolean(single(args[0]))
    return isFormulaError(flag) ? flag : !flag
  }),

  // These four inspect the argument instead of using it, so they receive the error.
  define(['ÉERROS', 'EERROS', 'ISERROR'], 1, 1, (args) => isFormulaError(single(args[0])), true),
  define(['É.NÃO.DISP', 'ENAODISP', 'ISNA'], 1, 1, (args) => single(args[0]) === FormulaError.NA, true),
  define(['ÉNÚM', 'ENUM', 'ISNUMBER'], 1, 1, (args) => typeof single(args[0]) === 'number', true),
  define(['ÉTEXTO', 'ETEXTO', 'ISTEXT'], 1, 1, (args) => typeof single(args[0]) === 'string', true),
  define(['ÉCÉL.VAZIA', 'ECELVAZIA', 'ISBLANK'], 1, 1, (args) => single(args[0]) === null, true),
]
