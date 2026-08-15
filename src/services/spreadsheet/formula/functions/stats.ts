/** Contagem, média e extremos. */

import { FormulaError, isFormulaError } from '../errors.js'
import { matchesCriteria } from '../values.js'
import { VARIADIC, define, numbersIn, rowsOf, single, valuesIn, type FunctionDefinition } from './kit.js'

export const STATS: readonly FunctionDefinition[] = [
  define(['MÉDIA', 'MEDIA', 'AVERAGE'], 1, VARIADIC, (args) => {
    const numbers = numbersIn(args)
    if (isFormulaError(numbers)) return numbers
    // Sem nenhum número não existe média. Devolver zero seria inventar um
    // resultado que passaria despercebido num relatório.
    if (numbers.length === 0) return FormulaError.Div0

    return numbers.reduce((total, value) => total + value, 0) / numbers.length
  }),

  define(['MÁXIMO', 'MAXIMO', 'MAX'], 1, VARIADIC, (args) => {
    const numbers = numbersIn(args)
    if (isFormulaError(numbers)) return numbers
    // O Excel devolve zero para intervalo sem números, e não erro.
    return numbers.length === 0 ? 0 : Math.max(...numbers)
  }),

  define(['MÍNIMO', 'MINIMO', 'MIN'], 1, VARIADIC, (args) => {
    const numbers = numbersIn(args)
    if (isFormulaError(numbers)) return numbers
    return numbers.length === 0 ? 0 : Math.min(...numbers)
  }),

  /** Conta **números**. Texto e vazio não entram. */
  define(['CONT.NÚM', 'CONT.NUM', 'COUNT'], 1, VARIADIC, (args) => {
    const numbers = numbersIn(args)
    return isFormulaError(numbers) ? numbers : numbers.length
  }),

  /** Conta o que **não** está vazio, inclusive texto e erro. */
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
        // Célula vazia não conta, nem quando o critério é "<>x": senão uma
        // coluna com dez mil linhas em branco daria dez mil ocorrências.
        if (value === null) continue
        if (matchesCriteria(value, criteria)) count++
      }
    }

    return count
  }),
]
