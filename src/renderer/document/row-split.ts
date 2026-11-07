import type { EditorView } from '@tiptap/pm/view'
import { measureLines } from './line-boxes.js'
import { REPEATED_HEADER_CLASS } from './extensions/pagination.js'

/**
 * Word breaks a table row between lines, unless it is `w:cantSplit`: each cell keeps the lines that
 * fit and moves the rest to the next sheet. A cut can fall in the middle of another cell's line;
 * that line goes along, and the whole rest of the row moves down by the same amount (`carried`),
 * so every cell starts the next sheet aligned.
 */
export interface RowSplit {
  /** Flow coordinates. */
  readonly at: number
  readonly carried: number
  /** Resolved only if the cut is chosen. */
  readonly resolve: () => ResolvedRowSplit | null
}

export interface ResolvedRowSplit {
  /** Where each moved cell gets its spacer: the first character of its first moved line. */
  readonly positions: readonly number[]
  /** Paper's cut in each cell, relative to its content; `null` when the cell ended above. */
  readonly cells: readonly (number | null)[]
  /** The moved line that opens the sheet, highest of all: the repeated header goes above it. */
  readonly first: number
}

/**
 * A cut is the midpoint between the font areas of two lines (`measureLines`), a little above the
 * line boxes' boundary, while a paragraph's last line ends at its border: a line in the next cell
 * ending this close to the cut is the line beside it.
 */
const LINE_EDGE_PX = 2

/**
 * A line of text, or a row of a nested table, which gets a spacer in each of its cells. Without
 * positions (an image, a row whose cell opens with a table), it only moves whole with the row.
 */
interface Unit {
  readonly top: number
  readonly bottom: number
  readonly positionsOf: (() => number[] | null) | null
}

interface RowGeometry {
  readonly view: EditorView
  /** The element's top in flow coordinates. */
  readonly flowTop: (node: HTMLElement) => number
  readonly height: (node: HTMLElement) => number
}

/** A cut that leaves no line of any cell above it is the break before the row, not inside it. */
export function rowSplitsOf(row: HTMLTableRowElement, geometry: RowGeometry): RowSplit[] {
  const cells = Array.from(row.cells)
  const units = cells.map((cell) => unitsOf(cell, geometry))
  const top = geometry.flowTop(row)
  const bottom = top + geometry.height(row)
  const candidates = [
    ...new Set(
      units.flatMap((list) => list.filter((unit) => unit.positionsOf !== null).map((unit) => unit.top)),
    ),
  ]
    .filter((at) => at > top + 0.5 && at < bottom - 0.5)
    .sort((left, right) => left - right)

  const splits: RowSplit[] = []
  for (const at of candidates) {
    const split = splitAt(at, units, cells, geometry.view)
    if (split !== null && (splits.at(-1)?.at ?? -Infinity) < at - 0.5) splits.push(split)
  }
  return splits
}

function splitAt(
  at: number,
  units: readonly (readonly Unit[])[],
  cells: readonly HTMLTableCellElement[],
  view: EditorView,
): RowSplit | null {
  let carried = 0
  let kept = false
  const moved: { readonly cell: number; readonly unit: Unit }[] = []
  for (const [cell, list] of units.entries()) {
    const index = list.findIndex((unit) => unit.bottom > at + LINE_EDGE_PX)
    if (index > 0) kept = true
    if (index === -1) {
      kept ||= list.length > 0
      continue
    }
    const unit = list[index]!
    if (unit.positionsOf === null) return null
    carried = Math.max(carried, at - unit.top)
    moved.push({ cell, unit })
  }
  if (!kept || moved.length === 0) return null
  return {
    at,
    carried,
    resolve: () => resolveSplit(moved, cells, view),
  }
}

function resolveSplit(
  moved: readonly { readonly cell: number; readonly unit: Unit }[],
  cells: readonly HTMLTableCellElement[],
  view: EditorView,
): ResolvedRowSplit | null {
  const positions: number[] = []
  const offsets: (number | null)[] = cells.map(() => null)
  let first: { readonly top: number; readonly position: number } | null = null
  for (const { cell, unit } of moved) {
    const found = unit.positionsOf?.() ?? null
    if (found === null || found.length === 0) return null
    positions.push(...found)
    offsets[cell] = cellOffsetOf(view, cells[cell]!, found[0]!)
    if (first === null || unit.top < first.top) first = { top: unit.top, position: found[0]! }
  }
  return first === null ? null : { positions, cells: offsets, first: first.position }
}

/**
 * At the start of a paragraph the cut goes before it, or paper would keep an empty piece of it at
 * the foot of the cell.
 */
function cellOffsetOf(view: EditorView, cell: HTMLTableCellElement, position: number): number {
  const { doc } = view.state
  const start = view.posAtDOM(cell, 0)
  const depth = doc.resolve(start).depth
  let at = position
  let resolved = doc.resolve(at)
  while (resolved.depth > depth && resolved.parentOffset === 0) {
    at = resolved.before()
    resolved = doc.resolve(at)
  }
  return at - start
}

function unitsOf(container: HTMLElement, geometry: RowGeometry): Unit[] {
  const units: Unit[] = []
  for (const child of Array.from(container.children)) {
    if (!(child instanceof HTMLElement) || child.classList.contains(REPEATED_HEADER_CLASS)) continue
    if (/^(P|H[1-6])$/.test(child.tagName)) {
      // The item's marker sits on its first line: a spacer there would leave it on the sheet above.
      const opensItem = container.tagName === 'LI' && child === container.firstElementChild
      units.push(...linesOf(child, geometry, !opensItem))
    } else if (child.tagName === 'UL' || child.tagName === 'OL' || child.tagName === 'LI') {
      units.push(...unitsOf(child, geometry))
    } else {
      units.push(...nestedRowsOf(child, geometry))
    }
  }
  return units
}

function linesOf(textblock: HTMLElement, geometry: RowGeometry, firstMoves: boolean): Unit[] {
  const top = geometry.flowTop(textblock)
  const height = geometry.height(textblock)
  if (height <= 0) return []
  const lines = measureLines(geometry.view, textblock)
  const tops = [top, ...lines.starts.map((start) => top + start)]
  const first = firstMoves ? () => listOf(safePosition(geometry.view, textblock)) : null
  return tops.map((lineTop, index) => ({
    top: lineTop,
    bottom: tops[index + 1] ?? top + height,
    positionsOf: index === 0 ? first : () => listOf(lines.positionOf(index - 1)),
  }))
}

/** A nested table breaks between its rows; anything else moves whole. */
function nestedRowsOf(child: HTMLElement, geometry: RowGeometry): Unit[] {
  const table = child instanceof HTMLTableElement ? child : child.querySelector(':scope > table')
  if (!(table instanceof HTMLTableElement) || table.rows.length === 0) {
    const top = geometry.flowTop(child)
    const height = geometry.height(child)
    return height > 0 ? [{ top, bottom: top + height, positionsOf: null }] : []
  }
  return Array.from(table.rows).map((row) => {
    const top = geometry.flowTop(row)
    return { top, bottom: top + geometry.height(row), positionsOf: () => rowStartsOf(row, geometry.view) }
  })
}

/** The first character of each cell, where its spacer goes; `null` if a cell opens with no text. */
function rowStartsOf(row: HTMLTableRowElement, view: EditorView): number[] | null {
  const positions: number[] = []
  for (const cell of Array.from(row.cells)) {
    const first = cell.firstElementChild
    if (!(first instanceof HTMLElement) || !/^(P|H[1-6])$/.test(first.tagName)) return null
    const position = safePosition(view, first)
    if (position === null) return null
    positions.push(position)
  }
  return positions
}

function listOf(position: number | null): number[] | null {
  return position === null ? null : [position]
}

function safePosition(view: EditorView, textblock: HTMLElement): number | null {
  try {
    return view.posAtDOM(textblock, 0)
  } catch {
    return null
  }
}
