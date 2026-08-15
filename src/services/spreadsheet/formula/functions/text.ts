/** Funções de texto. */

import { FormulaError, isFormulaError } from '../errors.js'
import { toText } from '../values.js'
import { VARIADIC, define, numberArg, textArg, valuesIn, type FunctionDefinition } from './kit.js'

/**
 * Posições de texto no Excel começam em **um**, não em zero.
 *
 * Converter na entrada, e não espalhar `- 1` pelas funções, evita o erro de um
 * caractere que aparece só na borda.
 */
function start(position: number): number {
  return Math.trunc(position) - 1
}

export const TEXT: readonly FunctionDefinition[] = [
  define(['CONCATENAR', 'CONCAT', 'CONCATENATE'], 1, VARIADIC, (args) => {
    let result = ''
    for (const value of valuesIn(args)) {
      const text = toText(value)
      if (isFormulaError(text)) return text
      result += text
    }
    return result
  }),

  define(['NÚM.CARACT', 'NUM.CARACT', 'LEN'], 1, 1, (args) => {
    const text = textArg(args[0])
    // Conta pontos de código, e não unidades UTF-16: um emoji é um caractere.
    return isFormulaError(text) ? text : [...text].length
  }),

  define(['ESQUERDA', 'LEFT'], 1, 2, (args) => {
    const text = textArg(args[0])
    if (isFormulaError(text)) return text
    const count = args.length > 1 ? numberArg(args[1]) : 1
    if (isFormulaError(count)) return count
    if (count < 0) return FormulaError.Value

    return [...text].slice(0, Math.trunc(count)).join('')
  }),

  define(['DIREITA', 'RIGHT'], 1, 2, (args) => {
    const text = textArg(args[0])
    if (isFormulaError(text)) return text
    const count = args.length > 1 ? numberArg(args[1]) : 1
    if (isFormulaError(count)) return count
    if (count < 0) return FormulaError.Value

    const letters = [...text]
    const wanted = Math.trunc(count)
    return wanted === 0 ? '' : letters.slice(Math.max(0, letters.length - wanted)).join('')
  }),

  define(['EXT.TEXTO', 'MID'], 3, 3, (args) => {
    const text = textArg(args[0])
    if (isFormulaError(text)) return text
    const from = numberArg(args[1])
    if (isFormulaError(from)) return from
    const count = numberArg(args[2])
    if (isFormulaError(count)) return count
    if (from < 1 || count < 0) return FormulaError.Value

    const at = start(from)
    return [...text].slice(at, at + Math.trunc(count)).join('')
  }),

  define(['MAIÚSCULA', 'MAIUSCULA', 'UPPER'], 1, 1, (args) => {
    const text = textArg(args[0])
    return isFormulaError(text) ? text : text.toUpperCase()
  }),

  define(['MINÚSCULA', 'MINUSCULA', 'LOWER'], 1, 1, (args) => {
    const text = textArg(args[0])
    return isFormulaError(text) ? text : text.toLowerCase()
  }),

  define(['ARRUMAR', 'TRIM'], 1, 1, (args) => {
    const text = textArg(args[0])
    // ARRUMAR não só apara as pontas: ela também reduz espaços do meio a um só.
    // É para isso que ela serve ao limpar dado colado de outro sistema.
    return isFormulaError(text) ? text : text.trim().replaceAll(/\s+/g, ' ')
  }),

  define(['SUBSTITUIR', 'SUBSTITUTE'], 3, 3, (args) => {
    const text = textArg(args[0])
    if (isFormulaError(text)) return text
    const from = textArg(args[1])
    if (isFormulaError(from)) return from
    const to = textArg(args[2])
    if (isFormulaError(to)) return to

    return from.length === 0 ? text : text.replaceAll(from, to)
  }),

  define(['PROCURAR', 'SEARCH'], 2, 3, (args) => {
    const needle = textArg(args[0])
    if (isFormulaError(needle)) return needle
    const haystack = textArg(args[1])
    if (isFormulaError(haystack)) return haystack
    const from = args.length > 2 ? numberArg(args[2]) : 1
    if (isFormulaError(from)) return from
    if (from < 1) return FormulaError.Value

    // PROCURAR não diferencia maiúsculas — é LOCALIZAR que diferencia.
    const at = haystack.toUpperCase().indexOf(needle.toUpperCase(), start(from))
    return at < 0 ? FormulaError.Value : at + 1
  }),

  define(['LOCALIZAR', 'FIND'], 2, 3, (args) => {
    const needle = textArg(args[0])
    if (isFormulaError(needle)) return needle
    const haystack = textArg(args[1])
    if (isFormulaError(haystack)) return haystack
    const from = args.length > 2 ? numberArg(args[2]) : 1
    if (isFormulaError(from)) return from
    if (from < 1) return FormulaError.Value

    const at = haystack.indexOf(needle, start(from))
    return at < 0 ? FormulaError.Value : at + 1
  }),

  define(['VALOR', 'VALUE'], 1, 1, (args) => {
    const text = textArg(args[0])
    if (isFormulaError(text)) return text

    // Passa pela mesma leitura brasileira que a digitação na célula usa.
    const number = Number(
      text.includes(',') ? text.replaceAll('.', '').replace(',', '.') : text.replace(/\s/g, ''),
    )
    return Number.isFinite(number) && text.trim().length > 0 ? number : FormulaError.Value
  }),
]
