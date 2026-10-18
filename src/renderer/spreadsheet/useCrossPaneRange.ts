import { useEffect } from 'react'
import type { Sheet } from '@services/spreadsheet/model.js'
import { normalizeRange, singleCell, type Range } from '@services/spreadsheet/edit.js'
import { FROZEN_COLUMNS_PANE, SHEET_ROWS_PANE, gridPositionOf, sheetColumnOf } from './grid-position.js'
import { TEXT_FIELD } from './useSheetUndo.js'

interface PaneCell {
  readonly x: number
  readonly y: number
}

interface SheetCell {
  readonly row: number
  readonly column: number
}

/** The part of RevoGrid's selection used here: one store per pane, each with its own range. */
interface PaneSelection {
  setRange(start: PaneCell, end: PaneCell): void
  clearFocus(): void
}

export interface SelectionConnector {
  readonly stores: Record<number, Record<number, PaneSelection | undefined> | undefined>
  readonly storesByType: Record<string, number | undefined>
  readonly storesXToType: Record<number, string | undefined>
  readonly focusedStore: { readonly cell: PaneCell; readonly position: PaneCell } | null
}

interface GridElement extends Element {
  getProviders(): Promise<{ selection?: SelectionConnector } | undefined>
}

const ARROW_STEPS: Readonly<Record<string, SheetCell>> = {
  ArrowUp: { row: -1, column: 0 },
  ArrowDown: { row: 1, column: 0 },
  ArrowLeft: { row: 0, column: -1 },
  ArrowRight: { row: 0, column: 1 },
}

/** `MouseEvent.button` of the left button, and its bit in `MouseEvent.buttons`. */
export const PRIMARY_BUTTON = 0
const PRIMARY_BUTTON_HELD = 1

/** A press on a cell, followed until the button is released. */
interface Drag {
  readonly anchor: SheetCell
  last: SheetCell
  /** From the first time the pointer goes over the divide, the range is this hook's to draw. */
  crossed: boolean
}

const inFrozenPane = (cell: SheetCell, frozenColumns: number): boolean => cell.column < frozenColumns
const sameCell = (a: SheetCell, b: SheetCell): boolean => a.row === b.row && a.column === b.column

function rangeBetween(anchor: SheetCell, extent: SheetCell): Range {
  return normalizeRange({
    fromRow: anchor.row,
    fromColumn: anchor.column,
    toRow: extent.row,
    toColumn: extent.column,
  })
}

/** The grid's focused cell is the anchor: the corner of the range that Shift does not move. */
function anchorOf(selection: SelectionConnector, frozenColumns: number): SheetCell | null {
  const focused = selection.focusedStore
  if (focused === null) return null
  const pane = selection.storesXToType[focused.position.x]
  return { row: focused.cell.y, column: sheetColumnOf(focused.cell.x, pane, frozenColumns) }
}

function extentOf(range: Range, anchor: SheetCell): SheetCell {
  return {
    row: range.fromRow === anchor.row ? range.toRow : range.fromRow,
    column: range.fromColumn === anchor.column ? range.toColumn : range.fromColumn,
  }
}

function stepped(extent: SheetCell, step: SheetCell, sheet: Sheet): SheetCell {
  const clamp = (value: number, count: number): number => Math.min(Math.max(value, 0), count - 1)
  return {
    row: clamp(extent.row + step.row, sheet.rowCount),
    column: clamp(extent.column + step.column, sheet.columnCount),
  }
}

/** Each pane draws the part of the range inside it; the grid by itself never draws across them. */
export function paint(selection: SelectionConnector, range: Range, sheet: Sheet): void {
  const panes = selection.stores[selection.storesByType[SHEET_ROWS_PANE] ?? -1] ?? {}
  for (const [x, type] of Object.entries(selection.storesXToType)) {
    const store = panes[Number(x)]
    if (store === undefined) continue
    const offset = sheetColumnOf(0, type, sheet.frozenColumns)
    const width = type === FROZEN_COLUMNS_PANE ? sheet.frozenColumns : sheet.columnCount - offset
    const from = Math.max(range.fromColumn - offset, 0)
    const to = Math.min(range.toColumn - offset, width - 1)
    if (from > to) store.clearFocus()
    else store.setRange({ x: from, y: range.fromRow }, { x: to, y: range.toRow })
  }
}

/** The grid's selection, found again each time a cell takes focus; returns the unsubscribe. */
export function followSelection(
  element: HTMLElement,
  found: (selection: SelectionConnector | undefined) => void,
): () => void {
  const connect = (): void => {
    const grid = element.querySelector<GridElement>('revo-grid')
    void grid?.getProviders().then((providers) => found(providers?.selection))
  }
  connect()
  element.addEventListener('afterfocus', connect)
  return () => element.removeEventListener('afterfocus', connect)
}

function isTyping(target: EventTarget | null): boolean {
  return target instanceof Element && target.closest(TEXT_FIELD) !== null
}

function arrowStepOf(event: KeyboardEvent): SheetCell | undefined {
  if (!event.shiftKey || event.ctrlKey || event.metaKey || event.altKey) return undefined
  return ARROW_STEPS[event.key]
}

/**
 * RevoGrid keeps one selection per pane and extends a range only inside the pane of its focused
 * cell: Shift+arrow stops at the frozen columns, and Shift+click past them restarts there. With
 * frozen columns, the extension is done here, from the anchor, as in any spreadsheet; always, not
 * only across the divide, because the grid applies an arrow 70 ms late and the two would race. A
 * drag with the mouse is left to the grid until it crosses the divide, where the grid stops.
 */
export function useCrossPaneRange(
  container: { readonly current: HTMLDivElement | null },
  current: { readonly current: Sheet },
  selected: { readonly current: Range },
  setRange: (range: Range) => void,
): void {
  useEffect(() => {
    const element = container.current
    if (element === null) return
    let selection: SelectionConnector | undefined
    let drag: Drag | null = null

    /** `true` when the extension is this hook's, and the grid must not see the event. */
    const extendTo = (target: (anchor: SheetCell) => SheetCell): boolean => {
      const sheet = current.current
      if (sheet.frozenColumns === 0) return false
      const anchor = selection === undefined ? null : anchorOf(selection, sheet.frozenColumns)
      if (selection === undefined || anchor === null) return false
      const range = rangeBetween(anchor, target(anchor))
      paint(selection, range, sheet)
      setRange(range)
      return true
    }

    const followKey = (event: KeyboardEvent): void => {
      const step = arrowStepOf(event)
      if (step === undefined || isTyping(event.target)) return
      const extended = extendTo((anchor) =>
        stepped(extentOf(selected.current, anchor), step, current.current),
      )
      if (!extended) return
      event.preventDefault()
      event.stopPropagation()
    }

    const followClick = (event: MouseEvent): void => {
      const position = event.button === PRIMARY_BUTTON ? gridPositionOf(event) : null
      if (position === null) return
      // Clicking the focused cell again, the grid announces nothing, though it drops the range.
      if (!event.shiftKey) setRange(singleCell(position.row, position.column))
      drag = event.shiftKey ? null : { anchor: position, last: position, crossed: false }
      if (!event.shiftKey || !extendTo(() => position)) return
      event.preventDefault()
      event.stopPropagation()
    }

    const followDrag = (event: MouseEvent): void => {
      const position = drag === null ? null : gridPositionOf(event)
      if (drag === null || position === null) return
      if ((event.buttons & PRIMARY_BUTTON_HELD) === 0) {
        drag = null
        return
      }
      const { frozenColumns } = current.current
      drag.crossed ||= inFrozenPane(position, frozenColumns) !== inFrozenPane(drag.anchor, frozenColumns)
      if (!drag.crossed) return
      event.stopPropagation()
      if (sameCell(position, drag.last)) return
      drag.last = position
      extendTo(() => position)
    }

    const endDrag = (event: MouseEvent): void => {
      if (drag?.crossed === true) event.stopPropagation()
      drag = null
    }

    const unfollow = followSelection(element, (found) => {
      selection = found
    })
    element.addEventListener('mousedown', followClick, { capture: true })
    // Before the grid's own listener, on `document`.
    window.addEventListener('keydown', followKey, { capture: true })
    window.addEventListener('mousemove', followDrag, { capture: true })
    window.addEventListener('mouseup', endDrag, { capture: true })
    return () => {
      unfollow()
      element.removeEventListener('mousedown', followClick, { capture: true })
      window.removeEventListener('keydown', followKey, { capture: true })
      window.removeEventListener('mousemove', followDrag, { capture: true })
      window.removeEventListener('mouseup', endDrag, { capture: true })
    }
  }, [container, current, selected, setRange])
}
