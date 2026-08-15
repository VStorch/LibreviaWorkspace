/**
 * Reescrita de referências.
 *
 * Duas coisas mexem nas referências de uma fórmula, e mexem de formas
 * diferentes:
 *
 * - **copiar** a fórmula para outra célula desloca as referências relativas e
 *   deixa as absolutas onde estão. É para isso que o `$` existe;
 * - **inserir ou excluir** linha ou coluna desloca **todas**, inclusive as
 *   absolutas — porque a célula apontada de fato mudou de lugar. Referência
 *   para uma célula excluída vira `#REF!`.
 *
 * Sem isso, inserir uma linha moveria os dados e deixaria `SOMA(A1:A3)`
 * apontando para onde os dados não estão mais: um total errado, sem aviso
 * nenhum. É o defeito que a Fase 5 deixou em aberto ao ganhar o menu de
 * inserir e excluir.
 *
 * A reescrita é feita sobre os **símbolos**, e não sobre a árvore: assim a
 * fórmula volta com o espaçamento, as maiúsculas e os parênteses que o usuário
 * digitou, e só as referências afetadas mudam.
 */

import { FormulaError } from './errors.js'
import { formatReference, parseReference, type CellRef } from './references.js'
import { TokenKind, tokenize, type Token } from './tokenize.js'

/** O que aconteceu com uma referência. */
type Moved = CellRef | 'broken'

/** Eixo da operação: as regras são as mesmas para linha e para coluna. */
const Axis = { Row: 'row', Column: 'column' } as const
type Axis = (typeof Axis)[keyof typeof Axis]

/**
 * Desloca as referências relativas de uma fórmula copiada.
 *
 * Sair da planilha pela esquerda ou por cima vira `#REF!`, como no Excel:
 * copiar `=A1` para a coluna A não tem para onde apontar.
 */
export function translateFormula(formula: string, rowDelta: number, columnDelta: number): string {
  return rewrite(formula, (ref) => {
    const row = ref.rowAbsolute ? ref.row : ref.row + rowDelta
    const column = ref.columnAbsolute ? ref.column : ref.column + columnDelta
    if (row < 0 || column < 0) return 'broken'

    return { ...ref, row, column }
  })
}

/**
 * Onde a linha ou coluna foi mexida, do ponto de vista da fórmula que está
 * sendo ajustada.
 *
 * As duas informações são necessárias porque uma referência sem nome de
 * planilha aponta para a planilha da **própria fórmula**: `=A1` numa fórmula da
 * aba "Resumo" não é afetada por uma linha inserida em "Dados", mas
 * `=Dados!A1` é — mesmo estando na mesma fórmula.
 */
export interface AdjustTarget {
  /** Nome da planilha onde a linha ou coluna foi inserida ou excluída. */
  readonly sheet: string
  /** A fórmula sendo ajustada mora nessa mesma planilha? */
  readonly own: boolean
}

/**
 * Ajusta as referências depois de inserir ou excluir linhas.
 *
 * `delta` positivo insere, negativo exclui.
 */
export function adjustForRows(formula: string, at: number, delta: number, target: AdjustTarget): string {
  return adjust(formula, Axis.Row, at, delta, target)
}

export function adjustForColumns(formula: string, at: number, delta: number, target: AdjustTarget): string {
  return adjust(formula, Axis.Column, at, delta, target)
}

/**
 * Troca o nome da planilha nas referências que a citam.
 *
 * Sem isso, renomear uma aba transformaria em `#REF!` toda fórmula que apontava
 * para ela — uma destruição em massa causada por um gesto que o usuário
 * considera cosmético.
 */
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
 * As duas pontas de um intervalo, movidas juntas.
 *
 * Juntas porque excluir parte de um intervalo o **encolhe**: `A1:A5` com as
 * três primeiras linhas excluídas vira `A1:A2`. Tratar as pontas
 * separadamente transformaria a primeira em `#REF!` e destruiria a fórmula
 * inteira por uma exclusão que o Excel absorve sem reclamar.
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

/**
 * Percorre os símbolos trocando só as referências.
 *
 * Um intervalo chega como `referência : referência`, e as duas pontas vão
 * juntas para `onRange` — quem trata cada ponta por si erra a exclusão parcial.
 */
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
