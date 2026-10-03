/** Guardar a árvore permite avaliar sem reanalisar, achar as dependências e reescrever referências. */

import type { FormulaError } from './errors.js'
import type { CellRef } from './references.js'

export type BinaryOperator = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '<=' | '>' | '>='

export type Node =
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly value: FormulaError }
  | { readonly kind: 'reference'; readonly ref: CellRef }
  | { readonly kind: 'range'; readonly from: CellRef; readonly to: CellRef }
  | { readonly kind: 'unary'; readonly operator: '-' | '+'; readonly operand: Node }
  | { readonly kind: 'percent'; readonly operand: Node }
  | {
      readonly kind: 'binary'
      readonly operator: BinaryOperator
      readonly left: Node
      readonly right: Node
    }
  | { readonly kind: 'call'; readonly name: string; readonly args: readonly Node[] }

export function* walk(node: Node): Generator<Node> {
  yield node

  switch (node.kind) {
    case 'unary':
    case 'percent':
      yield* walk(node.operand)
      break
    case 'binary':
      yield* walk(node.left)
      yield* walk(node.right)
      break
    case 'call':
      for (const arg of node.args) yield* walk(arg)
      break
    default:
      break
  }
}
