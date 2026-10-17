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

interface SelectionConnector {
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

/** `MouseEvent.button` of the left button. */
const PRIMARY_BUTTON = 0

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
function paint(selection: SelectionConnector, range: Range, sheet: Sheet): void {
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
 * only across the divide, because the grid applies an arrow 70 ms late and the two would race.
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

    const connect = (): void => {
      const grid = element.querySelector<GridElement>('revo-grid')
      void grid?.getProviders().then((providers) => {
        selection = providers?.selection
      })
    }

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
      if (!event.shiftKey || !extendTo(() => position)) return
      event.preventDefault()
      event.stopPropagation()
    }

    connect()
    element.addEventListener('afterfocus', connect)
    element.addEventListener('mousedown', followClick, { capture: true })
    // Before the grid's own listener, on `document`.
    window.addEventListener('keydown', followKey, { capture: true })
    return () => {
      element.removeEventListener('afterfocus', connect)
      element.removeEventListener('mousedown', followClick, { capture: true })
      window.removeEventListener('keydown', followKey, { capture: true })
    }
  }, [container, current, selected, setRange])
}
