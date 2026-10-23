/** Ascending precedence, with Excel's table: `=1+2*3` is seven. */

import type { MessageKey, Vars } from '@shared/i18n/index.js'
import type { BinaryOperator, Node } from './ast.js'
import { ParseError, isFormulaError, type FormulaError } from './errors.js'
import { parseReference, type CellRef } from './references.js'
import { TokenKind, tokenize, type Token } from './tokenize.js'

const PRECEDENCE: Record<BinaryOperator, number> = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '<=': 1,
  '>': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
}

/** `2^3^2` is `2^(3^2)`, not `(2^3)^2`. Only exponentiation is right-associative. */
const RIGHT_ASSOCIATIVE = new Set<BinaryOperator>(['^'])

/** Throws `ParseError` with a position, so the message says **where** the problem is. */
export function parseFormula(formula: string): Node {
  const text = formula.startsWith('=') ? formula.slice(1) : formula
  const tokens = tokenize(text)
  if (tokens.length === 0) throw new ParseError('spreadsheet.formula.empty', 0)

  const parser = new Parser(tokens, text.length)
  const node = parser.expression(0)
  parser.expectEnd()
  return node
}

/** Returns `null` when the formula does not parse. */
export function tryParseFormula(formula: string): Node | null {
  try {
    return parseFormula(formula)
  } catch {
    return null
  }
}

class Parser {
  #at = 0

  constructor(
    private readonly tokens: readonly Token[],
    private readonly end: number,
  ) {}

  expression(minimum: number): Node {
    let left = this.unary()

    for (;;) {
      const token = this.peek()
      if (token?.kind !== TokenKind.Operator) break

      const operator = token.text as BinaryOperator
      const precedence = PRECEDENCE[operator]
      if (precedence === undefined || precedence < minimum) break

      this.#at++
      const next = RIGHT_ASSOCIATIVE.has(operator) ? precedence : precedence + 1
      left = { kind: 'binary', operator, left, right: this.expression(next) }
    }

    return left
  }

  /** In Excel unary minus binds tighter than exponentiation: `=-2^2` is four. */
  unary(): Node {
    const token = this.peek()

    if (token?.kind === TokenKind.Operator && (token.text === '-' || token.text === '+')) {
      this.#at++
      const operand = this.unary()
      return { kind: 'unary', operator: token.text, operand }
    }

    return this.postfix(this.primary())
  }

  postfix(node: Node): Node {
    let result = node
    for (;;) {
      const token = this.peek()
      if (token?.kind !== TokenKind.Operator || token.text !== '%') return result
      this.#at++
      result = { kind: 'percent', operand: result }
    }
  }

  primary(): Node {
    const token = this.take()

    switch (token.kind) {
      case TokenKind.Number:
        return { kind: 'number', value: numberOf(token) }

      case TokenKind.Text:
        return { kind: 'text', value: token.text }

      case TokenKind.Boolean:
        return { kind: 'boolean', value: token.text === 'VERDADEIRO' || token.text === 'TRUE' }

      case TokenKind.Error:
        return { kind: 'error', value: asError(token) }

      case TokenKind.Open: {
        const inner = this.expression(0)
        this.expect(TokenKind.Close, 'spreadsheet.formula.unclosedParenthesis')
        return inner
      }

      case TokenKind.Name:
        return this.call(token)

      case TokenKind.Reference:
        return this.reference(token)

      default:
        throw new ParseError('spreadsheet.formula.unexpectedToken', token.position, { token: token.text })
    }
  }

  call(name: Token): Node {
    this.expect(TokenKind.Open, 'spreadsheet.formula.missingCallParenthesis', { name: name.text })

    const args: Node[] = []
    if (this.peek()?.kind !== TokenKind.Close) {
      for (;;) {
        args.push(this.expression(0))
        if (this.peek()?.kind !== TokenKind.Separator) break
        this.#at++
      }
    }

    this.expect(TokenKind.Close, 'spreadsheet.formula.unclosedCall', { name: name.text })
    return { kind: 'call', name: name.text, args }
  }

  reference(token: Token): Node {
    const from = this.cellRef(token)

    const next = this.peek()
    if (next?.kind !== TokenKind.Operator || next.text !== ':') return { kind: 'reference', ref: from }

    this.#at++
    const end = this.expect(TokenKind.Reference, 'spreadsheet.formula.rangeEndMissing')
    const to = this.cellRef(end)

    // A range lives on a single sheet: `Plan1!A1:Plan2!B2` is refused.
    if (to.sheet !== undefined && to.sheet !== from.sheet) {
      throw new ParseError('spreadsheet.formula.rangeAcrossSheets', end.position)
    }

    // Stored already ordered: `B4:A1` and `A1:B4` are the same rectangle. The sheet name stays only
    // on the first end, as Excel writes it.
    const low = corner(from, to, Math.min)
    return {
      kind: 'range',
      from: from.sheet === undefined ? low : { ...low, sheet: from.sheet },
      to: corner(from, to, Math.max),
    }
  }

  cellRef(token: Token): CellRef {
    const ref = parseReference(token.text)
    if (ref === null)
      throw new ParseError('spreadsheet.formula.invalidCell', token.position, { token: token.text })
    return ref
  }

  peek(): Token | undefined {
    return this.tokens[this.#at]
  }

  take(): Token {
    const token = this.tokens[this.#at]
    if (token === undefined) throw new ParseError('spreadsheet.formula.endedEarly', this.end)
    this.#at++
    return token
  }

  expect(kind: TokenKind, problem: MessageKey, vars?: Vars): Token {
    const token = this.peek()
    if (token?.kind !== kind) throw new ParseError(problem, token?.position ?? this.end, vars)
    this.#at++
    return token
  }

  expectEnd(): void {
    const token = this.peek()
    if (token !== undefined)
      throw new ParseError('spreadsheet.formula.trailingToken', token.position, { token: token.text })
  }
}

function numberOf(token: Token): number {
  const value = Number(token.text.replace(',', '.'))
  if (!Number.isFinite(value))
    throw new ParseError('spreadsheet.formula.invalidNumber', token.position, { token: token.text })
  return value
}

function asError(token: Token): FormulaError {
  if (!isFormulaError(token.text))
    throw new ParseError('spreadsheet.formula.unknownErrorValue', token.position, { token: token.text })
  return token.text
}

/**
 * The `$` follows the cell that ended up in that corner, otherwise copying the formula would
 * change.
 */
function corner(from: CellRef, to: CellRef, pick: (a: number, b: number) => number): CellRef {
  const row = pick(from.row, to.row) === from.row ? from : to
  const column = pick(from.column, to.column) === from.column ? from : to

  return {
    row: row.row,
    column: column.column,
    rowAbsolute: row.rowAbsolute,
    columnAbsolute: column.columnAbsolute,
  }
}
