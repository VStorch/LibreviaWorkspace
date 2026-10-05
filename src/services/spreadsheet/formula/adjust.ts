/**
 * Copying a formula shifts only relative references; inserting or deleting a row or column shifts
 * all of them, absolute ones included, and a reference to a deleted cell becomes `#REF!`. The
 * rewrite works on tokens, not on the tree, so the formula comes back as the user typed it.
 */

import { FormulaError } from './errors.js'
import { formatReference, parseReference, type CellRef } from './references.js'
import { TokenKind, tokenize, type Token } from './tokenize.js'

type Moved = CellRef | 'broken'

const Axis = { Row: 'row', Column: 'column' } as const
type Axis = (typeof Axis)[keyof typeof Axis]

/** Leaving the sheet to the left or the top becomes `#REF!`, as in Excel. */
export function translateFormula(formula: string, rowDelta: number, columnDelta: number): string {
  return rewrite(formula, (ref) => {
    const row = ref.rowAbsolute ? ref.row : ref.row + rowDelta
    const column = ref.columnAbsolute ? ref.column : ref.column + columnDelta
    if (row < 0 || column < 0) return 'broken'

    return { ...ref, row, column }
  })
}

/** A reference without a sheet name points to the **formula's own** sheet. */
export interface AdjustTarget {
  readonly sheet: string
  /** Does the formula being adjusted live on that same sheet? */
  readonly own: boolean
}

/** A positive `delta` inserts, a negative one deletes. */
export function adjustForRows(formula: string, at: number, delta: number, target: AdjustTarget): string {
  return adjust(formula, Axis.Row, at, delta, target)
}

export function adjustForColumns(formula: string, at: number, delta: number, target: AdjustTarget): string {
  return adjust(formula, Axis.Column, at, delta, target)
}

/** Otherwise renaming a tab would turn every formula pointing to it into `#REF!`. */
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

/** A lone position: gone if it was in the deleted band. */
function movePoint(position: number, at: number, delta: number): number | null {
  if (position < at) return position
  if (delta > 0) return position + delta

  const removed = -delta
  // Inside the deleted band there is nowhere left to point.
  return position < at + removed ? null : position + delta
}

/**
 * Together, because deleting part of a range **shrinks** it: `A1:A5` without its first three rows
 * becomes `A1:A2`, as in Excel.
 */
function moveSpan(from: number, to: number, at: number, delta: number): { from: number; to: number } | null {
  if (delta > 0) {
    // Inserting inside the range stretches it; inserting after it leaves it alone.
    return { from: from >= at ? from + delta : from, to: to >= at ? to + delta : to }
  }

  const removed = -delta
  const after = at + removed

  const start = from >= after ? from + delta : from >= at ? at : from
  const end = to >= after ? to + delta : to >= at ? at - 1 : to

  // The whole range fell in the deleted band.
  return end < start ? null : { from: start, to: end }
}

/** Both ends of a range go together to `onRange`. */
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
    // A formula that cannot even be read has no reference to adjust.
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

    // Without range handling, each end moves on its own, which is right for copying, where the
    // shift is the same on both sides.
    const moved = onSingle(ref)
    result += body.slice(cursor, token.position)
    result += textOf(moved)
    cursor = token.position + token.text.length
  }

  return prefix + result + body.slice(cursor)
}

/** The second end of a range never repeats the sheet name. */
function textOf(moved: Moved, dropSheet = false): string {
  if (moved === 'broken') return FormulaError.Ref
  return formatReference(dropSheet ? { ...moved, sheet: undefined } : moved)
}
