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
}

/**
 * A cut is the midpoint between the font areas of two lines (`measureLines`), a little above the
 * line boxes' boundary, while a paragraph's last line ends at its border: a line in the next cell
 * ending this close to the cut is the line beside it.
 */
const LINE_EDGE_PX = 2

/** A line of text, or a nested table, which only moves whole and has nowhere to get a spacer. */
interface Unit {
  readonly top: number
  readonly bottom: number
  readonly positionOf: (() => number | null) | null
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
      units.flatMap((list) => list.filter((unit) => unit.positionOf !== null).map((unit) => unit.top)),
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
    if (unit.positionOf === null) return null
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
  for (const { cell, unit } of moved) {
    const position = unit.positionOf?.() ?? null
    if (position === null) return null
    positions.push(position)
    offsets[cell] = cellOffsetOf(view, cells[cell]!, position)
  }
  return { positions, cells: offsets }
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
      const top = geometry.flowTop(child)
      if (geometry.height(child) > 0)
        units.push({ top, bottom: top + geometry.height(child), positionOf: null })
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
  const first = firstMoves ? () => safePosition(geometry.view, textblock) : null
  return tops.map((lineTop, index) => ({
    top: lineTop,
    bottom: tops[index + 1] ?? top + height,
    positionOf: index === 0 ? first : () => lines.positionOf(index - 1),
  }))
}

function safePosition(view: EditorView, textblock: HTMLElement): number | null {
  try {
    return view.posAtDOM(textblock, 0)
  } catch {
    return null
  }
}
