/**
 * Copiar uma fórmula desloca só as referências relativas; inserir ou excluir
 * linha ou coluna desloca todas, inclusive as absolutas, e a referência para
 * uma célula excluída vira `#REF!`. A reescrita é feita sobre os símbolos, e não
 * sobre a árvore, para a fórmula voltar como a pessoa a digitou.
 */

import { FormulaError } from './errors.js'
import { formatReference, parseReference, type CellRef } from './references.js'
import { TokenKind, tokenize, type Token } from './tokenize.js'

type Moved = CellRef | 'broken'

const Axis = { Row: 'row', Column: 'column' } as const
type Axis = (typeof Axis)[keyof typeof Axis]

/** Sair da planilha pela esquerda ou por cima vira `#REF!`, como no Excel. */
export function translateFormula(formula: string, rowDelta: number, columnDelta: number): string {
  return rewrite(formula, (ref) => {
    const row = ref.rowAbsolute ? ref.row : ref.row + rowDelta
    const column = ref.columnAbsolute ? ref.column : ref.column + columnDelta
    if (row < 0 || column < 0) return 'broken'

    return { ...ref, row, column }
  })
}

/** Referência sem nome de planilha aponta para a planilha da **própria fórmula**. */
export interface AdjustTarget {
  readonly sheet: string
  /** A fórmula sendo ajustada mora nessa mesma planilha? */
  readonly own: boolean
}

/** `delta` positivo insere, negativo exclui. */
export function adjustForRows(formula: string, at: number, delta: number, target: AdjustTarget): string {
  return adjust(formula, Axis.Row, at, delta, target)
}

export function adjustForColumns(formula: string, at: number, delta: number, target: AdjustTarget): string {
  return adjust(formula, Axis.Column, at, delta, target)
}

/** Sem isto, renomear uma aba transformaria em `#REF!` toda fórmula que aponta para ela. */
export function renameSheetInFormula(formula: string, from: string, to: string): string {
  const target = from.toUpperCase()

  return rewrite(formula, (ref) =>
    ref.sheet !== undefined && ref.sheet.toUpperCase() === target ? { ...ref, sheet: to } : ref,
  )
}

function adjust(formula: string, axis: Axis, at: number, delta: number, target: AdjustTarget): string {
  const affects = (ref: CellRef): boolean =>
    ref.sheet === undefined ? target.own : ref.sheet.toUpperCase() === target.sheet.toUpperCase()

  return rewrite(
    formula,
    (ref) => {
      if (!affects(ref)) return ref

      const moved = movePoint(ref[axis], at, delta)
      return moved === null ? 'broken' : { ...ref, [axis]: moved }
    },
    (from, to) => {
      if (!affects(from)) return { from, to }

      const ends = moveSpan(from[axis], to[axis], at, delta)
      if (ends === null) return { from: 'broken', to: 'broken' }

      return { from: { ...from, [axis]: ends.from }, to: { ...to, [axis]: ends.to } }
    },
  )
}

/** Uma posição sozinha: some se estava na faixa excluída. */
function movePoint(position: number, at: number, delta: number): number | null {
  if (position < at) return position
  if (delta > 0) return position + delta

  const removed = -delta
  // Dentro da faixa excluída não sobra para onde apontar.
  return position < at + removed ? null : position + delta
}

/**
 * Juntas, porque excluir parte de um intervalo o **encolhe**: `A1:A5` sem as
 * três primeiras linhas vira `A1:A2`, como no Excel.
 */
function moveSpan(from: number, to: number, at: number, delta: number): { from: number; to: number } | null {
  if (delta > 0) {
    // Inserir dentro do intervalo o estica; inserir depois não o toca.
    return { from: from >= at ? from + delta : from, to: to >= at ? to + delta : to }
  }

  const removed = -delta
  const after = at + removed

  const start = from >= after ? from + delta : from >= at ? at : from
  const end = to >= after ? to + delta : to >= at ? at - 1 : to

  // O intervalo inteiro caiu na faixa excluída.
  return end < start ? null : { from: start, to: end }
}

/** As duas pontas de um intervalo vão juntas para `onRange`. */
function rewrite(
  formula: string,
  onSingle: (ref: CellRef) => Moved,
  onRange?: (from: CellRef, to: CellRef) => { from: Moved; to: Moved },
): string {
  const prefix = formula.startsWith('=') ? '=' : ''
  const body = formula.slice(prefix.length)

  let tokens: Token[]
  try {
    tokens = tokenize(body)
  } catch {
    // Fórmula que nem chega a ser lida não tem referência para ajustar.
    return formula
  }

  let result = ''
  let cursor = 0

  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index]!
    if (token.kind !== TokenKind.Reference) continue

    const ref = parseReference(token.text)
    if (ref === null) continue

    const colon = tokens[index + 1]
    const second = tokens[index + 2]
    const isRange =
      colon?.kind === TokenKind.Operator &&
      colon.text === ':' &&
      second?.kind === TokenKind.Reference &&
      parseReference(second.text) !== null

    if (isRange && onRange !== undefined) {
      const to = parseReference(second.text)!
      const moved = onRange(ref, to)

      result += body.slice(cursor, token.position)
      result += textOf(moved.from) + ':' + textOf(moved.to, true)
      cursor = second.position + second.text.length
      index += 2
      continue
    }

    // Sem tratamento de intervalo, cada ponta anda por si — que é o certo para
    // a cópia, onde o deslocamento é o mesmo dos dois lados.
    const moved = onSingle(ref)
    result += body.slice(cursor, token.position)
    result += textOf(moved)
    cursor = token.position + token.text.length
  }

  return prefix + result + body.slice(cursor)
}

/** A segunda ponta de um intervalo nunca repete o nome da planilha. */
function textOf(moved: Moved, dropSheet = false): string {
  if (moved === 'broken') return FormulaError.Ref
  return formatReference(dropSheet ? { ...moved, sheet: undefined } : moved)
}
