/**
 * XLSX stores English function names, commas between arguments and decimal points; the app shows
 * Excel's Portuguese dialect. Translation is **character by character**, not through the tree: the
 * surgical save only rewrites a cell whose formula changed, and rebuilding would normalize spaces
 * and case.
 */

import { localizedName } from './functions/index.js'

/** Errors written inside the formula, as in `=SEERRO(A1;#N/D)`. */
const ERROR_PAIRS: readonly (readonly [string, string])[] = [
  ['#DIV/0!', '#DIV/0!'],
  ['#VALOR!', '#VALUE!'],
  ['#REF!', '#REF!'],
  ['#NOME?', '#NAME?'],
  ['#NÚM!', '#NUM!'],
  ['#N/D', '#N/A'],
]

interface Dialect {
  readonly language: 'pt' | 'en'
  /** Decimal separators accepted in the source. */
  readonly decimalsIn: string
  /** Separador decimal a escrever. */
  readonly decimalOut: string
  /** The source's, then the destination's. */
  readonly separator: readonly [string, string]
  readonly errors: ReadonlyMap<string, string>
}

const TO_APP: Dialect = {
  language: 'pt',
  // In the file the comma always separates arguments and is never decimal.
  decimalsIn: '.',
  decimalOut: ',',
  separator: [',', ';'],
  errors: new Map(ERROR_PAIRS.map(([portuguese, english]) => [english, portuguese])),
}

const TO_FILE: Dialect = {
  language: 'en',
  // Both, because the dot is what comes from pasting a foreign spreadsheet.
  decimalsIn: '.,',
  decimalOut: '.',
  separator: [';', ','],
  errors: new Map(ERROR_PAIRS),
}

/** Read from XLSX, in the app's dialect. */
export function fromXlsxFormula(formula: string): string {
  return convert(formula, TO_APP)
}

/** From the app, in the file's dialect. */
export function toXlsxFormula(formula: string): string {
  return convert(formula, TO_FILE)
}

function convert(formula: string, dialect: Dialect): string {
  let out = ''
  let at = 0

  while (at < formula.length) {
    const char = formula[at]!

    // Quoted text and apostrophe-quoted sheet names pass intact: a comma inside them is content,
    // not a separator.
    if (char === '"' || char === "'") {
      const end = closingAt(formula, at, char)
      out += formula.slice(at, end)
      at = end
      continue
    }

    // A literal array `{1,2;3,4}`: the app does not understand it, and it passes intact.
    if (char === '{') {
      const end = formula.indexOf('}', at)
      const stop = end === -1 ? formula.length : end + 1
      out += formula.slice(at, stop)
      at = stop
      continue
    }

    if (char === '#') {
      const literal = longestMatch(formula, at, dialect.errors)
      if (literal !== undefined) {
        out += dialect.errors.get(literal)!
        at += literal.length
        continue
      }
    }

    if (isDigit(char)) {
      const number = readNumber(formula, at, dialect.decimalsIn)
      out += rewriteDecimal(number, dialect)
      at += number.length
      continue
    }

    if (char === dialect.separator[0]) {
      out += dialect.separator[1]
      at++
      continue
    }

    const word = readWord(formula, at)
    if (word.length > 0) {
      // Only a function with the parenthesis attached, as in the parser: a cell named `SOMA` is not
      // one.
      out += formula[at + word.length] === '(' ? localizedName(word, dialect.language) : word
      at += word.length
      continue
    }

    out += char
    at++
  }

  return out
}

/**
 * Already past the closing quote. A doubled quote (`""`) does not close, and the apostrophe follows
 * the same convention.
 */
function closingAt(formula: string, start: number, quote: string): number {
  let at = start + 1
  while (at < formula.length) {
    if (formula[at] === quote) {
      if (formula[at + 1] === quote) {
        at += 2
        continue
      }
      return at + 1
    }
    at++
  }
  // An unclosed quote: copy the rest instead of inventing a closing.
  return formula.length
}

function longestMatch(formula: string, at: number, table: ReadonlyMap<string, string>): string | undefined {
  let best: string | undefined
  for (const candidate of table.keys()) {
    if (!formula.startsWith(candidate, at)) continue
    if (best === undefined || candidate.length > best.length) best = candidate
  }
  return best
}

function isDigit(char: string | undefined): boolean {
  return char !== undefined && char >= '0' && char <= '9'
}

/** Integer or decimal, with exponent. Mirrors the parser's. */
function readNumber(formula: string, start: number, decimals: string): string {
  let at = start
  let seenSeparator = false

  while (at < formula.length) {
    const char = formula[at]!
    if (isDigit(char)) {
      at++
      continue
    }
    if (decimals.includes(char) && !seenSeparator && isDigit(formula[at + 1])) {
      seenSeparator = true
      at += 2
      continue
    }
    if ((char === 'e' || char === 'E') && at > start) {
      if (isDigit(formula[at + 1])) {
        at += 2
        continue
      }
      if ((formula[at + 1] === '+' || formula[at + 1] === '-') && isDigit(formula[at + 2])) {
        at += 3
        continue
      }
    }
    break
  }

  return formula.slice(start, at)
}

function rewriteDecimal(number: string, dialect: Dialect): string {
  for (const separator of dialect.decimalsIn) {
    const dot = number.indexOf(separator)
    if (dot !== -1) {
      return number.slice(0, dot) + dialect.decimalOut + number.slice(dot + 1)
    }
  }
  return number
}

/** The same set as the parser: `$`, `!` and `.` belong to `$A$1`, `Plan1!A1` and `CONT.NÚM`. */
function readWord(formula: string, start: number): string {
  let at = start
  while (at < formula.length && /[\p{L}\p{N}_$.!]/u.test(formula[at]!)) at++
  return formula.slice(start, at)
}
