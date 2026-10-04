/**
 * Como no Excel em português, a vírgula é decimal e o ponto e vírgula separa
 * argumentos: aceitar os dois papéis tornaria `SOMA(1,5)` ambíguo. O ponto
 * também é decimal, porque é o que sai ao colar de planilha estrangeira.
 */

import { ParseError } from './errors.js'

export const TokenKind = {
  Number: 'number',
  Text: 'text',
  Boolean: 'boolean',
  /** Referência de célula, com ou sem `$` e com ou sem nome de planilha. */
  Reference: 'reference',
  /** Nome de função, sempre seguido de `(`. */
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

/** Operadores de dois caracteres primeiro: `<=` não pode virar `<` e `=`. */
const OPERATORS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', ':'] as const

const TRUE_WORDS = new Set(['VERDADEIRO', 'TRUE'])
const FALSE_WORDS = new Set(['FALSO', 'FALSE'])

/** Erros que o usuário pode digitar literalmente, como `=SEERRO(A1;#N/D)`. */
const ERROR_LITERALS = ['#DIV/0!', '#VALOR!', '#REF!', '#NOME?', '#NÚM!', '#N/D', '#CIRC!']

const WHITESPACE: ReadonlySet<string> = new Set([' ', '\t', '\n', '\r'])

const PUNCTUATION: ReadonlyMap<string, TokenKind> = new Map([
  ['(', TokenKind.Open],
  [')', TokenKind.Close],
  [';', TokenKind.Separator],
])

interface ReadToken {
  readonly token: Token
  /** O texto da fórmula consumido, que difere do token nas aspas dobradas e nas maiúsculas. */
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

  // O número vem antes da referência: um dígito nunca começa referência, e
  // ler ao contrário faria `1e3` virar `1` seguido de `e3`.
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
    // A vírgula solta é quase sempre um separador de argumentos: a dica diz isso.
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
  // Sobrou referência. Se não for uma, o analisador reclama com posição.
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
    // Expoente: `1e3`, `2E-5`. O `e` só conta se vier dígito depois.
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

/** `""` vale uma aspa, como no Excel. Devolve o conteúdo já sem as aspas. */
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

/** ``, `Planilha1!A1` e `CONT.NÚM` são uma palavra só; o nome entre apóstrofos entra inteiro. */
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
