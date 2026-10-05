/**
 * The dependency graph is walked in topological order, so `=B1+1` never reads B1's old value.
 * Cycles are detected in the same walk.
 */

import type { Cell, Sheet, WorkbookModel } from '../model.js'
import { walk, type Node } from './ast.js'
import { FormulaError } from './errors.js'
import { evaluate, type EvalContext } from './evaluate.js'
import { tryParseFormula } from './parse.js'
import type { CellRef } from './references.js'
import type { Scalar } from './values.js'

interface FormulaCell {
  readonly sheet: number
  readonly row: number
  readonly column: number
  /** As stored in the map, to write the result back. */
  readonly ref: string
  readonly node: Node | null
}

const key = (sheet: number, row: number, column: number): string => `${sheet}|${row}|${column}`

/**
 * A numeric key, not `"A1"`: `SOMA(A1:A10000)` does ten thousand reads. The column limit is
 * Excel's, so keys do not collide.
 */
const COLUMN_SPAN = 16_384
const at = (row: number, column: number): number => row * COLUMN_SPAN + column

/** A sheet without formulas comes back as the same object: React compares by identity. */
export function recalculate(workbook: WorkbookModel, now: () => Date = () => new Date()): WorkbookModel {
  const cells = collect(workbook)
  if (cells.length === 0) return workbook

  const byName = sheetsByName(workbook)
  const index = workbook.sheets.map(indexOf)
  const results: { cell: FormulaCell; value: Scalar }[] = []

  for (const cell of order(cells, byName)) {
    // A formula that does not parse only arrives from a hand-edited file: the UI refuses it before
    // saving.
    const value =
      cell.node === null
        ? FormulaError.Value
        : // The context is built per cell because a reference without a sheet name points to the
          // **formula's** sheet, which changes with each one.
          evaluate(cell.node, contextFor(cell.sheet, index, byName, now))

    // Writing into the index is what makes the topological order hold: whatever comes later and
    // depends on this cell already reads the new value.
    index[cell.sheet]?.set(at(cell.row, cell.column), value)
    results.push({ cell, value })
  }

  return apply(workbook, results)
}

function indexOf(sheet: Sheet): Map<number, Scalar> {
  const values = new Map<number, Scalar>()

  for (const [ref, cell] of Object.entries(sheet.cells)) {
    const position = positionOf(ref)
    if (position === null || cell.value === undefined) continue
    values.set(at(position.row, position.column), cell.value)
  }

  return values
}

function sheetsByName(workbook: WorkbookModel): ReadonlyMap<string, number> {
  const byName = new Map<string, number>()
  // Excel ignores case in sheet names, and the user types `plan1!A1` expecting it to work.
  for (const [index, sheet] of workbook.sheets.entries()) byName.set(sheet.name.toUpperCase(), index)
  return byName
}

function collect(workbook: WorkbookModel): FormulaCell[] {
  const cells: FormulaCell[] = []

  for (const [sheet, model] of workbook.sheets.entries()) {
    for (const [ref, cell] of Object.entries(model.cells)) {
      if (cell.formula === undefined) continue

      const position = positionOf(ref)
      if (position === null) continue

      cells.push({ sheet, ...position, ref, node: tryParseFormula(cell.formula) })
    }
  }

  return cells
}

/** Reads from the index, already updated by the formulas computed before this one. */
function contextFor(
  own: number,
  index: readonly Map<number, Scalar>[],
  byName: ReadonlyMap<string, number>,
  now: () => Date,
): EvalContext {
  return {
    now,
    valueAt: (ref: CellRef): Scalar => {
      const sheet = sheetIndexOf(ref, byName, own)
      // A sheet name that does not exist: the tab was deleted or renamed.
      if (sheet === null) return FormulaError.Ref

      return index[sheet]?.get(at(ref.row, ref.column)) ?? null
    },
  }
}

/** Without a name, it is the formula's own sheet, which changes with each cell. */
function sheetIndexOf(ref: CellRef, byName: ReadonlyMap<string, number>, fallback: number): number | null {
  if (ref.sheet === undefined) return fallback
  return byName.get(ref.sheet.toUpperCase()) ?? null
}

/** An explicit stack, not recursion: ten thousand chained formulas would overflow the stack. */
function order(cells: readonly FormulaCell[], byName: ReadonlyMap<string, number>): FormulaCell[] {
  const byKey = new Map<string, FormulaCell>()
  for (const cell of cells) byKey.set(key(cell.sheet, cell.row, cell.column), cell)

  const dependencies = new Map<string, string[]>()
  for (const cell of cells) {
    dependencies.set(key(cell.sheet, cell.row, cell.column), dependenciesOf(cell, byKey, byName))
  }

  const sorted: FormulaCell[] = []
  const state = new Map<string, 'visiting' | 'done'>()
  const circular = new Set<string>()

  for (const start of byKey.keys()) {
    if (state.has(start)) continue

    // Each frame records which dependency it stopped at, to resume from there.
    const stack: { at: string; next: number }[] = [{ at: start, next: 0 }]
    state.set(start, 'visiting')

    while (stack.length > 0) {
      const frame = stack[stack.length - 1]!
      const deps = dependencies.get(frame.at) ?? []

      if (frame.next >= deps.length) {
        stack.pop()
        state.set(frame.at, 'done')
        const cell = byKey.get(frame.at)
        if (cell !== undefined && !circular.has(frame.at)) sorted.push(cell)
        continue
      }

      const dependency = deps[frame.next++]!
      const seen = state.get(dependency)

      if (seen === 'done') continue
      if (seen === 'visiting') {
        // The cycle closed: everything on the stack from it on takes part.
        const from = stack.findIndex((entry) => entry.at === dependency)
        for (const entry of stack.slice(from)) circular.add(entry.at)
        continue
      }

      state.set(dependency, 'visiting')
      stack.push({ at: dependency, next: 0 })
    }
  }

  // Circular cells go first, already with the error: their dependents inherit it through normal
  // propagation instead of reading an old value.
  const broken: FormulaCell[] = []
  for (const at of circular) {
    const cell = byKey.get(at)
    if (cell !== undefined) broken.push({ ...cell, node: { kind: 'error', value: FormulaError.Circular } })
  }

  return [...broken, ...sorted]
}

/**
 * Only formulas impose order. Ranges are crossed against the formula list, otherwise
 * `SOMA(A1:A10000)` would cost ten thousand steps.
 */
function dependenciesOf(
  cell: FormulaCell,
  byKey: ReadonlyMap<string, FormulaCell>,
  byName: ReadonlyMap<string, number>,
): string[] {
  if (cell.node === null) return []

  const found = new Set<string>()

  for (const node of walk(cell.node)) {
    if (node.kind === 'reference') {
      const sheet = sheetIndexOf(node.ref, byName, cell.sheet)
      if (sheet === null) continue

      const at = key(sheet, node.ref.row, node.ref.column)
      if (byKey.has(at)) found.add(at)
      continue
    }

    if (node.kind !== 'range') continue

    const sheet = sheetIndexOf(node.from, byName, cell.sheet)
    if (sheet === null) continue

    for (const other of byKey.values()) {
      if (other.sheet !== sheet) continue
      if (other.row < node.from.row || other.row > node.to.row) continue
      if (other.column < node.from.column || other.column > node.to.column) continue
      found.add(key(other.sheet, other.row, other.column))
    }
  }

  // A formula referring to itself directly is already a cycle.
  found.delete(key(cell.sheet, cell.row, cell.column))
  if (referencesItself(cell, byName)) found.add(key(cell.sheet, cell.row, cell.column))

  return [...found]
}

function referencesItself(cell: FormulaCell, byName: ReadonlyMap<string, number>): boolean {
  if (cell.node === null) return false

  for (const node of walk(cell.node)) {
    if (node.kind === 'reference') {
      const sheet = sheetIndexOf(node.ref, byName, cell.sheet)
      if (sheet === cell.sheet && node.ref.row === cell.row && node.ref.column === cell.column) return true
      continue
    }
    if (node.kind !== 'range') continue

    const sheet = sheetIndexOf(node.from, byName, cell.sheet)
    if (sheet !== cell.sheet) continue
    if (cell.row < node.from.row || cell.row > node.to.row) continue
    if (cell.column < node.from.column || cell.column > node.to.column) continue
    return true
  }

  return false
}

/** A sheet with no changed value comes back as the same object, and so does the workbook. */
function apply(
  workbook: WorkbookModel,
  results: readonly { cell: FormulaCell; value: Scalar }[],
): WorkbookModel {
  const changed = new Map<number, Record<string, Cell>>()

  for (const { cell, value } of results) {
    const sheet = workbook.sheets[cell.sheet]
    const previous = sheet?.cells[cell.ref]
    if (sheet === undefined || previous === undefined) continue

    // As in Excel, `=A1` over a blank cell is zero. Empty would drop the formula from the sparse
    // map on the next save.
    const calculated = value === null ? 0 : value
    if (calculated === previous.value) continue

    let cells = changed.get(cell.sheet)
    if (cells === undefined) {
      cells = { ...sheet.cells }
      changed.set(cell.sheet, cells)
    }
    cells[cell.ref] = { ...previous, value: calculated }
  }

  if (changed.size === 0) return workbook

  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet, index) => {
      const cells = changed.get(index)
      return cells === undefined ? sheet : { ...sheet, cells }
    }),
  }
}

function positionOf(ref: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)([0-9]+)$/.exec(ref)
  if (match === null) return null

  let column = 0
  for (const letter of match[1]!) column = column * 26 + (letter.charCodeAt(0) - 64)

  return { row: Number.parseInt(match[2]!, 10) - 1, column: column - 1 }
}
