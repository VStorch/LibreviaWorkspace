import { FormulaError, isFormulaError } from '../errors.js'
import { matchesCriteria } from '../values.js'
import { VARIADIC, define, numbersIn, rowsOf, single, valuesIn, type FunctionDefinition } from './kit.js'

export const STATS: readonly FunctionDefinition[] = [
  define(['MÉDIA', 'MEDIA', 'AVERAGE'], 1, VARIADIC, (args) => {
    const numbers = numbersIn(args)
    if (isFormulaError(numbers)) return numbers
    // Without numbers there is no average; zero would be an invented result.
    if (numbers.length === 0) return FormulaError.Div0

    return numbers.reduce((total, value) => total + value, 0) / numbers.length
  }),

  define(['MÁXIMO', 'MAXIMO', 'MAX'], 1, VARIADIC, (args) => {
    const numbers = numbersIn(args)
    if (isFormulaError(numbers)) return numbers
    // Excel returns zero for a range without numbers, not an error.
    return numbers.length === 0 ? 0 : Math.max(...numbers)
  }),

  define(['MÍNIMO', 'MINIMO', 'MIN'], 1, VARIADIC, (args) => {
    const numbers = numbersIn(args)
    if (isFormulaError(numbers)) return numbers
    return numbers.length === 0 ? 0 : Math.min(...numbers)
  }),

  /** Counts **numbers**. Text and empty cells are excluded. */
  define(['CONT.NÚM', 'CONT.NUM', 'COUNT'], 1, VARIADIC, (args) => {
    const numbers = numbersIn(args)
    return isFormulaError(numbers) ? numbers : numbers.length
  }),

  /** Counts what is **not** empty, text and errors included. */
  define(['CONT.VALORES', 'COUNTA'], 1, VARIADIC, (args) => {
    return valuesIn(args).filter((value) => value !== null && value !== '').length
  }),

  define(['CONTAR.VAZIO', 'COUNTBLANK'], 1, VARIADIC, (args) => {
    return valuesIn(args).filter((value) => value === null || value === '').length
  }),

  define(['CONT.SE', 'COUNTIF'], 2, 2, (args) => {
    const tested = rowsOf(args[0])
    if (isFormulaError(tested)) return tested
    const criteria = single(args[1])
    if (isFormulaError(criteria)) return criteria

    let count = 0
    for (const line of tested) {
      for (const value of line) {
        // An empty cell does not count, not even with the "<>x" criterion.
        if (value === null) continue
        if (matchesCriteria(value, criteria)) count++
      }
    }

    return count
  }),
]
