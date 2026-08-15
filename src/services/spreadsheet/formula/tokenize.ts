/**
 * Fórmula em texto → sequência de símbolos.
 *
 * Duas escolhas de idioma valem para todo o motor, e são as do Excel em
 * português:
 *
 * - **vírgula é decimal**: `=SOMA(1,5;2)` soma um e meio com dois;
 * - **ponto e vírgula separa argumentos**, justamente porque a vírgula já está
 *   ocupada. Aceitar os dois papéis para a vírgula tornaria `SOMA(1,5)`
 *   ambíguo — um argumento ou dois? — e a ambiguidade cairia sempre em cima de
 *   quem digitou um número decimal.
 *
 * O ponto também é aceito como decimal (`1.5`), porque não disputa papel com
 * nada e é o que sai ao colar de planilha estrangeira.
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

export function tokenize(formula: string): Token[] {
  const tokens: Token[] = []
  let at = 0

  while (at < formula.length) {
    const char = formula[at]!

    if (char === ' ' || char === '\t' || char === '\n' || char === '\r') {
      at++
      continue
    }

    if (char === '(') {
      tokens.push({ kind: TokenKind.Open, text: '(', position: at++ })
      continue
    }
    if (char === ')') {
      tokens.push({ kind: TokenKind.Close, text: ')', position: at++ })
      continue
    }
    if (char === ';') {
      tokens.push({ kind: TokenKind.Separator, text: ';', position: at++ })
      continue
    }

    if (char === '"') {
      const token = readText(formula, at)
      tokens.push(token)
      at += token.text.length + quotesIn(token.text) + 2
      continue
    }

    if (char === '#') {
      const literal = ERROR_LITERALS.find((error) => formula.startsWith(error, at))
      if (literal === undefined) throw new ParseError(`Erro desconhecido em "${formula.slice(at)}".`, at)
      tokens.push({ kind: TokenKind.Error, text: literal, position: at })
      at += literal.length
      continue
    }

    // O número vem antes da referência: um dígito nunca começa referência, e
    // ler ao contrário faria `1e3` virar `1` seguido de `e3`.
    if (isDigit(char) || ((char === ',' || char === '.') && isDigit(formula[at + 1]))) {
      const text = readNumber(formula, at)
      tokens.push({ kind: TokenKind.Number, text, position: at })
      at += text.length
      continue
    }

    const operator = OPERATORS.find((candidate) => formula.startsWith(candidate, at))
    if (operator !== undefined) {
      tokens.push({ kind: TokenKind.Operator, text: operator, position: at })
      at += operator.length
      continue
    }

    const word = readWord(formula, at)
    if (word.length === 0) {
      // A vírgula solta é quase sempre alguém separando argumentos com ela.
      // Dizer só "não entendi" mandaria o usuário procurar o erro errado.
      const hint =
        char === ',' ? 'Use ponto e vírgula para separar argumentos: SOMA(A1;B1).' : `Não entendi "${char}".`
      throw new ParseError(hint, at)
    }

    const upper = word.toUpperCase()
    if (TRUE_WORDS.has(upper) || FALSE_WORDS.has(upper)) {
      tokens.push({ kind: TokenKind.Boolean, text: upper, position: at })
    } else if (formula[at + word.length] === '(') {
      tokens.push({ kind: TokenKind.Name, text: upper, position: at })
    } else {
      // Sobrou referência. Se não for uma, o analisador reclama com posição.
      tokens.push({ kind: TokenKind.Reference, text: word, position: at })
    }
    at += word.length
    continue
  }

  return tokens
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

/**
 * Texto entre aspas, com `""` valendo uma aspa — a convenção do Excel.
 *
 * Devolve o conteúdo já sem as aspas: quem chama soma o que consumiu usando o
 * número de aspas duplicadas.
 */
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

/**
 * Palavra: nome de função ou referência.
 *
 * Aceita `$`, `!` e `.` porque `$A$1`, `Planilha1!A1` e `CONT.NÚM` são uma
 * palavra só. O nome de planilha entre apóstrofos entra inteiro, com espaços.
 */
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
