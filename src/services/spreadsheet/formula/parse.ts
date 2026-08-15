/**
 * Símbolos → árvore.
 *
 * Analisador de precedência ascendente: uma tabela de precedência em vez de uma
 * função por nível. A tabela é a do Excel, e é ela que faz `=1+2*3` valer sete e
 * `=2^3^2` valer quinhentos e doze — potência associa à direita.
 */

import type { BinaryOperator, Node } from './ast.js'
import { ParseError, isFormulaError, type FormulaError } from './errors.js'
import { parseReference, type CellRef } from './references.js'
import { TokenKind, tokenize, type Token } from './tokenize.js'

/** Precedência dos operadores binários. Maior liga mais forte. */
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

/** `2^3^2` é `2^(3^2)`, e não `(2^3)^2`. Só a potência associa à direita. */
const RIGHT_ASSOCIATIVE = new Set<BinaryOperator>(['^'])

/**
 * Analisa uma fórmula, com ou sem o `=` inicial.
 *
 * Lança `ParseError` com posição: a mensagem precisa dizer **onde** está o
 * problema, senão o usuário fica procurando parêntese numa fórmula longa.
 */
export function parseFormula(formula: string): Node {
  const text = formula.startsWith('=') ? formula.slice(1) : formula
  const tokens = tokenize(text)
  if (tokens.length === 0) throw new ParseError('A fórmula está vazia.', 0)

  const parser = new Parser(tokens, text.length)
  const node = parser.expression(0)
  parser.expectEnd()
  return node
}

/** Analisa sem lançar: devolve `null` quando a fórmula não fecha. */
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

  /**
   * Unário e sufixo `%`.
   *
   * No Excel o menos unário liga **mais forte** que a potência: `=-2^2` vale
   * quatro, e não menos quatro como na matemática. A regra é estranha, mas é a
   * que a planilha ao lado usa — e o resultado precisa bater com ela.
   */
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
        this.expect(TokenKind.Close, 'Faltou fechar um parêntese.')
        return inner
      }

      case TokenKind.Name:
        return this.call(token)

      case TokenKind.Reference:
        return this.reference(token)

      default:
        throw new ParseError(`Não esperava "${token.text}" aqui.`, token.position)
    }
  }

  call(name: Token): Node {
    this.expect(TokenKind.Open, `Faltou o parêntese depois de ${name.text}.`)

    const args: Node[] = []
    if (this.peek()?.kind !== TokenKind.Close) {
      for (;;) {
        args.push(this.expression(0))
        if (this.peek()?.kind !== TokenKind.Separator) break
        this.#at++
      }
    }

    this.expect(TokenKind.Close, `Faltou fechar o parêntese de ${name.text}.`)
    return { kind: 'call', name: name.text, args }
  }

  /** Uma referência, ou duas ligadas por `:` formando um intervalo. */
  reference(token: Token): Node {
    const from = this.cellRef(token)

    const next = this.peek()
    if (next?.kind !== TokenKind.Operator || next.text !== ':') return { kind: 'reference', ref: from }

    this.#at++
    const end = this.expect(TokenKind.Reference, 'Depois de ":" falta uma célula.')
    const to = this.cellRef(end)

    // Um intervalo vive numa planilha só. Aceitar `Plan1!A1:Plan2!B2` obrigaria
    // a inventar o que ele significa; recusar diz ao usuário o que houve.
    if (to.sheet !== undefined && to.sheet !== from.sheet) {
      throw new ParseError('Um intervalo não pode atravessar duas planilhas.', end.position)
    }

    // O intervalo é guardado já ordenado: `B4:A1` e `A1:B4` são o mesmo
    // retângulo, e quem lê a árvore não deveria precisar saber disso. O nome da
    // planilha fica só na primeira ponta, como o Excel escreve.
    const low = corner(from, to, Math.min)
    return {
      kind: 'range',
      from: from.sheet === undefined ? low : { ...low, sheet: from.sheet },
      to: corner(from, to, Math.max),
    }
  }

  cellRef(token: Token): CellRef {
    const ref = parseReference(token.text)
    if (ref === null) throw new ParseError(`"${token.text}" não é uma célula válida.`, token.position)
    return ref
  }

  peek(): Token | undefined {
    return this.tokens[this.#at]
  }

  take(): Token {
    const token = this.tokens[this.#at]
    if (token === undefined) throw new ParseError('A fórmula terminou antes do esperado.', this.end)
    this.#at++
    return token
  }

  expect(kind: TokenKind, message: string): Token {
    const token = this.peek()
    if (token?.kind !== kind) throw new ParseError(message, token?.position ?? this.end)
    this.#at++
    return token
  }

  expectEnd(): void {
    const token = this.peek()
    if (token !== undefined) throw new ParseError(`Sobrou "${token.text}" no fim da fórmula.`, token.position)
  }
}

function numberOf(token: Token): number {
  const value = Number(token.text.replace(',', '.'))
  if (!Number.isFinite(value)) throw new ParseError(`"${token.text}" não é um número.`, token.position)
  return value
}

function asError(token: Token): FormulaError {
  if (!isFormulaError(token.text))
    throw new ParseError(`"${token.text}" não é um erro conhecido.`, token.position)
  return token.text
}

/**
 * Canto do retângulo.
 *
 * O `$` acompanha a célula que ficou naquele canto: trocar a ordem de `B4:A1`
 * sem levar o `$` junto mudaria o que acontece ao copiar a fórmula.
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
