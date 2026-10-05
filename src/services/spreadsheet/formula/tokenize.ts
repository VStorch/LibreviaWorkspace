/**
 * As in Portuguese Excel, the comma is decimal and the semicolon separates arguments: accepting
 * both roles would make `SOMA(1,5)` ambiguous. The dot is also decimal, because that is what comes
 * from pasting a foreign spreadsheet.
 */

import { ParseError } from './errors.js'

export const TokenKind = {
  Number: 'number',
  Text: 'text',
  Boolean: 'boolean',
  /** With or without `$` and with or without a sheet name. */
  Reference: 'reference',
  /** Always followed by `(`. */
  Name: 'name',
  Operator: 'operator',
  Open: 'open',
  Close: 'close',
  Separator: 'separator',
  Error: 'error',
} as const
export type TokenKind = (typeof TokenKind)[keyof typeof TokenKind]

export interface Token {
  readonly kind: TokenKind
  readonly text: string
  readonly position: number
}

/** Two-character operators first: `<=` must not become `<` and `=`. */
const OPERATORS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', ':'] as const

const TRUE_WORDS = new Set(['VERDADEIRO', 'TRUE'])
const FALSE_WORDS = new Set(['FALSO', 'FALSE'])

/** Errors the user may type literally, as in `=SEERRO(A1;#N/D)`. */
const ERROR_LITERALS = ['#DIV/0!', '#VALOR!', '#REF!', '#NOME?', '#NÚM!', '#N/D', '#CIRC!']

const WHITESPACE: ReadonlySet<string> = new Set([' ', '\t', '\n', '\r'])

const PUNCTUATION: ReadonlyMap<string, TokenKind> = new Map([
  ['(', TokenKind.Open],
  [')', TokenKind.Close],
  [';', TokenKind.Separator],
])

interface ReadToken {
  readonly token: Token
  /** The consumed formula text, which differs from the token in doubled quotes and case. */
  readonly length: number
}

export function tokenize(formula: string): Token[] {
  const tokens: Token[] = []
  let at = 0

  while (at < formula.length) {
    if (WHITESPACE.has(formula[at]!)) {
      at++
      continue
    }
    const { token, length } = readToken(formula, at)
    tokens.push(token)
    at += length
  }

  return tokens
}

function readToken(formula: string, at: number): ReadToken {
  const char = formula[at]!

  const punctuation = PUNCTUATION.get(char)
  if (punctuation !== undefined) return whole({ kind: punctuation, text: char, position: at })

  if (char === '"') {
    const token = readText(formula, at)
    return { token, length: token.text.length + quotesIn(token.text) + 2 }
  }

  if (char === '#') {
    const literal = ERROR_LITERALS.find((error) => formula.startsWith(error, at))
    if (literal === undefined) throw new ParseError(`Erro desconhecido em "${formula.slice(at)}".`, at)
    return whole({ kind: TokenKind.Error, text: literal, position: at })
  }

  // Numbers come before references: a digit never starts a reference, and reading the other way
  // would turn `1e3` into `1` followed by `e3`.
  if (isDigit(char) || ((char === ',' || char === '.') && isDigit(formula[at + 1]))) {
    return whole({ kind: TokenKind.Number, text: readNumber(formula, at), position: at })
  }

  const operator = OPERATORS.find((candidate) => formula.startsWith(candidate, at))
  if (operator !== undefined) return whole({ kind: TokenKind.Operator, text: operator, position: at })

  return readWordToken(formula, at)
}

function whole(token: Token): ReadToken {
  return { token, length: token.text.length }
}

function readWordToken(formula: string, at: number): ReadToken {
  const char = formula[at]!
  const word = readWord(formula, at)
  if (word.length === 0) {
    // A loose comma is almost always an argument separator: the hint says so.
    const hint =
      char === ',' ? 'Use ponto e vírgula para separar argumentos: SOMA(A1;B1).' : `Não entendi "${char}".`
    throw new ParseError(hint, at)
  }

  const upper = word.toUpperCase()
  if (TRUE_WORDS.has(upper) || FALSE_WORDS.has(upper)) {
    return { token: { kind: TokenKind.Boolean, text: upper, position: at }, length: word.length }
  }
  if (formula[at + word.length] === '(') {
    return { token: { kind: TokenKind.Name, text: upper, position: at }, length: word.length }
  }
  // A reference is what is left. If it is not one, the parser complains with a position.
  return whole({ kind: TokenKind.Reference, text: word, position: at })
}

function isDigit(char: string | undefined): boolean {
  return char !== undefined && char >= '0' && char <= '9'
}

function readNumber(formula: string, start: number): string {
  let at = start
  let seenSeparator = false

  while (at < formula.length) {
    const char = formula[at]!
    if (isDigit(char)) {
      at++
      continue
    }
    if ((char === ',' || char === '.') && !seenSeparator && isDigit(formula[at + 1])) {
      seenSeparator = true
      at += 2
      continue
    }
    // Exponent: `1e3`, `2E-5`. The `e` only counts if a digit follows.
    if ((char === 'e' || char === 'E') && at > start) {
      const next = formula[at + 1]
      const afterSign = formula[at + 2]
      if (isDigit(next)) {
        at += 2
        continue
      }
      if ((next === '+' || next === '-') && isDigit(afterSign)) {
        at += 3
        continue
      }
    }
    break
  }

  return formula.slice(start, at)
}

/** `""` is one quote, as in Excel. Returns the content without the quotes. */
function readText(formula: string, start: number): Token {
  let at = start + 1
  let value = ''

  while (at < formula.length) {
    const char = formula[at]!
    if (char === '"') {
      if (formula[at + 1] === '"') {
        value += '"'
        at += 2
        continue
      }
      return { kind: TokenKind.Text, text: value, position: start }
    }
    value += char
    at++
  }

  throw new ParseError('Faltou fechar as aspas do texto.', start)
}

function quotesIn(text: string): number {
  let count = 0
  for (const char of text) if (char === '"') count++
  return count
}

/** `$A$1`, `Planilha1!A1` and `CONT.NÚM` are a single word; a name in apostrophes enters whole. */
function readWord(formula: string, start: number): string {
  let at = start

  if (formula[at] === "'") {
    at++
    while (at < formula.length && formula[at] !== "'") at++
    if (at >= formula.length) throw new ParseError('Faltou fechar o apóstrofo do nome da planilha.', start)
    at++
    if (formula[at] !== '!') throw new ParseError('Depois do nome da planilha falta o "!".', start)
    at++
  }

  while (at < formula.length && isWordChar(formula[at]!)) at++
  return formula.slice(start, at)
}

function isWordChar(char: string): boolean {
  return /[\p{L}\p{N}_$.!]/u.test(char)
}
