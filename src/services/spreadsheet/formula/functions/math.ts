import { FormulaError, isFormulaError } from '../errors.js'
import { matchesCriteria, type Scalar } from '../values.js'
import { VARIADIC, define, numberArg, numbersIn, rowsOf, single, type FunctionDefinition } from './kit.js'

/** `.5` rounds away from zero, as in Excel; `Math.round` would take −2.5 to −2. */
function roundHalfAway(value: number, digits: number): number {
  const factor = 10 ** digits
  const scaled = value * factor
  // Without the correction, 1.005 would become 1.00 through binary.
  const corrected = Number(scaled.toPrecision(15))
  return (corrected < 0 ? -Math.round(-corrected) : Math.round(corrected)) / factor
}

function toDigits(value: number | FormulaError): number | FormulaError {
  return isFormulaError(value) ? value : Math.trunc(value)
}

export const MATH: readonly FunctionDefinition[] = [
  define(['SOMA', 'SUM'], 1, VARIADIC, (args) => {
    const numbers = numbersIn(args)
    return isFormulaError(numbers) ? numbers : numbers.reduce((total, value) => total + value, 0)
  }),

  define(['ARRED', 'ROUND'], 2, 2, (args) => {
    const value = numberArg(args[0])
    if (isFormulaError(value)) return value
    const digits = toDigits(numberArg(args[1]))
    return isFormulaError(digits) ? digits : roundHalfAway(value, digits)
  }),

  define(['ARREDONDAR.PARA.CIMA', 'ROUNDUP'], 2, 2, (args) => {
    const value = numberArg(args[0])
    if (isFormulaError(value)) return value
    const digits = toDigits(numberArg(args[1]))
    if (isFormulaError(digits)) return digits

    const factor = 10 ** digits
    return (value < 0 ? -Math.ceil(-value * factor) : Math.ceil(value * factor)) / factor
  }),

  define(['ARREDONDAR.PARA.BAIXO', 'ROUNDDOWN'], 2, 2, (args) => {
    const value = numberArg(args[0])
    if (isFormulaError(value)) return value
    const digits = toDigits(numberArg(args[1]))
    if (isFormulaError(digits)) return digits

    const factor = 10 ** digits
    return (value < 0 ? -Math.floor(-value * factor) : Math.floor(value * factor)) / factor
  }),

  define(['ABS'], 1, 1, (args) => {
    const value = numberArg(args[0])
    return isFormulaError(value) ? value : Math.abs(value)
  }),

  define(['INT'], 1, 1, (args) => {
    const value = numberArg(args[0])
    // INT really rounds down: INT(-2.5) is -3, not -2.
    return isFormulaError(value) ? value : Math.floor(value)
  }),

  define(['TRUNCAR', 'TRUNC'], 1, 2, (args) => {
    const value = numberArg(args[0])
    if (isFormulaError(value)) return value
    const digits = args.length > 1 ? toDigits(numberArg(args[1])) : 0
    if (isFormulaError(digits)) return digits

    const factor = 10 ** digits
    return Math.trunc(value * factor) / factor
  }),

  define(['RESTO', 'MOD'], 2, 2, (args) => {
    const value = numberArg(args[0])
    if (isFormulaError(value)) return value
    const divisor = numberArg(args[1])
    if (isFormulaError(divisor)) return divisor
    if (divisor === 0) return FormulaError.Div0

    // The remainder follows the divisor's sign, as in Excel, not the dividend's, as `%` does.
    return value - divisor * Math.floor(value / divisor)
  }),

  define(['RAIZ', 'SQRT'], 1, 1, (args) => {
    const value = numberArg(args[0])
    if (isFormulaError(value)) return value
    return value < 0 ? FormulaError.Num : Math.sqrt(value)
  }),

  define(['POTÊNCIA', 'POTENCIA', 'POWER'], 2, 2, (args) => {
    const base = numberArg(args[0])
    if (isFormulaError(base)) return base
    const exponent = numberArg(args[1])
    if (isFormulaError(exponent)) return exponent

    const result = base ** exponent
    if (Number.isNaN(result)) return FormulaError.Num
    return Number.isFinite(result) ? result : FormulaError.Num
  }),

  /** The sum range is shifted from the first one's corner, as in Excel. */
  define(['SOMASE', 'SUMIF'], 2, 3, (args) => {
    const tested = rowsOf(args[0])
    if (isFormulaError(tested)) return tested
    const criteria = single(args[1])
    if (isFormulaError(criteria)) return criteria

    const summed = args.length > 2 ? rowsOf(args[2]) : tested
    if (isFormulaError(summed)) return summed

    let total = 0
    for (let row = 0; row < tested.length; row++) {
      const line = tested[row]!
      for (let column = 0; column < line.length; column++) {
        if (!matchesCriteria(line[column] ?? null, criteria)) continue

        const value: Scalar = summed[row]?.[column] ?? null
        if (isFormulaError(value)) return value
        if (typeof value === 'number') total += value
      }
    }

    return total
  }),
]
