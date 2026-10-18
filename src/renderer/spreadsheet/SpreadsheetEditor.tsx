import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { RevoGrid, type ColumnRegular } from '@revolist/react-datagrid'
import type {
  AfterEditEvent,
  BeforeSaveDataDetails,
  ChangedRange,
  FocusAfterRenderEvent,
  RangeArea,
  RevoGridCustomEvent,
} from '@revolist/revogrid'
import { formatCell } from '@services/spreadsheet/format.js'
import { DEFAULT_COLUMN_WIDTH, columnName, getCell, type Sheet } from '@services/spreadsheet/model.js'
import {
  editableText,
  normalizeRange,
  rangeContains,
  singleCell,
  toggleStyle,
  writeText,
  type Range,
} from '@services/spreadsheet/edit.js'
import { fillRange } from '@services/spreadsheet/fill.js'
import type { StructuralChange } from '@services/spreadsheet/structure.js'
import { FormulaBar } from './FormulaBar.js'
import { usePreferences } from '../state/preferences.js'
import { SpreadsheetToolbar } from './SpreadsheetToolbar.js'
import { SheetContextMenu, type MenuPosition } from './SheetContextMenu.js'
import { cellStyleOf } from './cell-style.js'
import { SHEET_ROWS_PANE, gridPositionOf, sheetColumnOf } from './grid-position.js'
import { useCrossPaneRange } from './useCrossPaneRange.js'
import { useSheetClipboard } from './useSheetClipboard.js'
import { useFormatShortcuts } from './useFormatShortcuts.js'
import { useTypeAhead } from './useTypeAhead.js'

/**
 * The grid works with object rows, and the model is a sparse map by A1 reference: the translation
 * lives here.
 */

type GridRow = Record<string, string>

/** `MouseEvent.button` of the right button. */
const SECONDARY_BUTTON = 2

export function SpreadsheetEditor({
  sheet,
  onChange,
  onStructure,
  onRange,
  readOnly = false,
}: {
  sheet: Sheet
  onChange: (sheet: Sheet) => void
  /** A **workbook** operation: a row inserted here changes `=Dados!A5` written on another tab. */
  onStructure: (change: StructuralChange) => void
  /** The selection, for what lives beside the grid: the assistant works on it. */
  onRange?: (range: Range) => void
  /** Reading and scrolling keep working. */
  readOnly?: boolean
}): React.JSX.Element {
  const showToolbar = usePreferences((state) => state.preferences.showToolbar)
  // So the handler does not capture an old state.
  const current = useRef(sheet)
  current.current = sheet

  // Here, not in the grid: the toolbar needs it.
  const [range, setRange] = useState<Range>(() => singleCell(0, 0))
  const selection = useRef(range)
  selection.current = range
  useEffect(() => onRange?.(range), [onRange, range])

  const typeAhead = useTypeAhead(readOnly)

  const { applyChange, applyStructure } = useGuardedWrites(onChange, onStructure, readOnly)

  const { columns, source } = useGridModel(sheet, current)
  const { handleEdit, handleEditStart } = useEditHandlers(current, applyChange, typeAhead)
  const container = useRef<HTMLDivElement>(null)
  useGridRange(container, current, setRange)
  useCrossPaneRange(container, current, selection, setRange)
  useGridTakesFocus(container)
  useSheetClipboard(current, selection, applyChange)
  const { handleResize, handleFocus, handleAutofill } = useSelectionHandlers(
    current,
    applyChange,
    typeAhead,
    setRange,
  )
  const { menu, handleContextMenu, closeMenu } = useSheetMenu(container, selection, setRange)

  useFormatShortcuts((key) => applyChange(toggleStyle(current.current, selection.current, key)))

  return (
    <div ref={container} className="sheet" onContextMenu={handleContextMenu}>
      {showToolbar && <SpreadsheetToolbar sheet={sheet} range={range} onChange={applyChange} />}

      <FormulaBar
        sheet={sheet}
        range={range}
        onCommit={(text) => applyChange(writeText(current.current, range.fromRow, range.fromColumn, text))}
      />

      {menu !== null && (
        <SheetContextMenu
          sheet={sheet}
          range={range}
          position={menu}
          onChange={onChange}
          onStructure={applyStructure}
          onClose={closeMenu}
        />
      )}

      <RevoGrid
        columns={columns}
        source={source}
        theme="compact"
        resize={true}
        range={true}
        readonly={readOnly}
        rowHeaders={true}
        useClipboard={true}
        rowDefinitions={Array.from({ length: sheet.frozenRows }, (_, index) => ({
          type: 'rowPinStart' as const,
          index,
          size: 24,
        }))}
        onAfteredit={handleEdit}
        onBeforeeditstart={handleEditStart}
        onAftercolumnresize={handleResize}
        onAfterfocus={handleFocus}
        onBeforeautofill={handleAutofill}
      />
    </div>
  )
}

type SheetRef = { readonly current: Sheet }
type TypeAhead = ReturnType<typeof useTypeAhead>

function writeAt(sheet: Sheet, row: number, prop: string, raw: unknown): Sheet {
  const column = Number.parseInt(prop.slice(1), 10)
  if (!Number.isInteger(column)) return sheet

  return writeText(sheet, row, column, typeof raw === 'string' ? raw : String(raw ?? ''))
}

function useGridModel(sheet: Sheet, current: SheetRef): { columns: ColumnRegular[]; source: GridRow[] } {
  const columns = useMemo<ColumnRegular[]>(
    () =>
      Array.from({ length: sheet.columnCount }, (_, index) => {
        const column: ColumnRegular = {
          prop: `c${index}`,
          name: columnName(index),
          size: sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH,
          resizable: true,
          // The style lives in the model: it survives saving and scrolling, which recreates the
          // cells.
          cellProperties: ({ rowIndex }) => ({
            style: cellStyleOf(current.current, rowIndex, index),
          }),
        }
        if (index < sheet.frozenColumns) column.pin = 'colPinStart'
        return column
      }),
    [sheet],
  )

  /** One row per visible position, generated on demand from the sparse map. */
  const source = useMemo<GridRow[]>(() => {
    const rows: GridRow[] = []
    for (let row = 0; row < sheet.rowCount; row++) {
      const cells: GridRow = {}
      for (let column = 0; column < sheet.columnCount; column++) {
        cells[`c${column}`] = formatCell(getCell(sheet, row, column))
      }
      rows.push(cells)
    }
    return rows
  }, [sheet])

  return { columns, source }
}

function useEditHandlers(
  current: SheetRef,
  applyChange: (sheet: Sheet) => void,
  typeAhead: TypeAhead,
): {
  handleEdit: (event: RevoGridCustomEvent<AfterEditEvent>) => void
  handleEditStart: (event: RevoGridCustomEvent<BeforeSaveDataDetails>) => void
} {
  /** An edited cell or a pasted range, which arrives in another shape. */
  const handleEdit = useCallback(
    (event: RevoGridCustomEvent<AfterEditEvent>) => {
      const detail = event.detail
      let updated = current.current

      // `newRange` only exists when pasting; both declare `data`.
      if ('newRange' in detail) {
        for (const [row, values] of Object.entries(detail.data)) {
          for (const [prop, value] of Object.entries(values as Record<string, unknown>)) {
            updated = writeAt(updated, Number(row), prop, value)
          }
        }
      } else {
        updated = writeAt(updated, detail.rowIndex, String(detail.prop), detail.val)
      }

      applyChange(updated)

      // Pasting and confirming with the mouse do not go through Enter.
      typeAhead.begin()
    },
    [applyChange, typeAhead],
  )

  /**
   * Shows the **formula**: leaving the cell by accident would write the result over it. A text that
   * reads as a formula shows with its apostrophe, so Enter keeps it text.
   */
  const handleEditStart = useCallback((event: RevoGridCustomEvent<BeforeSaveDataDetails>) => {
    const detail = event.detail
    const column = Number.parseInt(String(detail.prop).slice(1), 10)
    if (!Number.isInteger(column)) return

    const cell = getCell(current.current, detail.rowIndex, column)
    const text = editableText(cell)
    const reinterpreted = typeof cell?.value === 'string' && text !== cell.value
    if (cell?.formula !== undefined || reinterpreted) detail.val = text
  }, [])

  return { handleEdit, handleEditStart }
}

function useSelectionHandlers(
  current: SheetRef,
  applyChange: (sheet: Sheet) => void,
  typeAhead: TypeAhead,
  setRange: (range: Range) => void,
): {
  handleResize: (event: CustomEvent<Record<number, ColumnRegular>>) => void
  handleFocus: (event: RevoGridCustomEvent<FocusAfterRenderEvent>) => void
  handleAutofill: (event: RevoGridCustomEvent<ChangedRange>) => void
} {
  const handleResize = useCallback(
    (event: CustomEvent<Record<number, ColumnRegular>>) => {
      const widths = { ...current.current.columnWidths }
      for (const [index, column] of Object.entries(event.detail)) {
        if (column.size !== undefined) widths[Number(index)] = column.size
      }
      applyChange({ ...current.current, columnWidths: widths })
    },
    [applyChange],
  )

  const handleFocus = useCallback(
    (event: RevoGridCustomEvent<FocusAfterRenderEvent>) => {
      const { rowIndex, colIndex, colType } = event.detail
      if (!Number.isInteger(rowIndex) || !Number.isInteger(colIndex)) return

      setRange(singleCell(rowIndex, sheetColumnOf(colIndex, colType, current.current.frozenColumns)))
      typeAhead.settle()
    },
    [typeAhead],
  )

  /** On the model: the grid would copy the displayed text, which for a formula is the result. */
  const handleAutofill = useCallback(
    (event: RevoGridCustomEvent<ChangedRange>) => {
      const { oldRange, newRange } = event.detail
      if (oldRange === null || newRange === null) return

      event.preventDefault()
      applyChange(
        fillRange(
          current.current,
          { fromRow: oldRange.y, fromColumn: oldRange.x, toRow: oldRange.y1, toColumn: oldRange.x1 },
          { fromRow: newRange.y, fromColumn: newRange.x, toRow: newRange.y1, toColumn: newRange.x1 },
        ),
      )
    },
    [applyChange],
  )

  return { handleResize, handleFocus, handleAutofill }
}

/**
 * The grid has no public event for a range drawn with the mouse or Shift+arrows (`beforerange` is
 * only for autofill, and `getSelectedRange` lags 70 ms behind the arrow keys). Its selection layer
 * announces each range with `setrange`, in the coordinates of its own pane: that is why the RevoGrid
 * version is pinned.
 */
function useGridRange(
  container: { readonly current: HTMLDivElement | null },
  current: SheetRef,
  setRange: (range: Range) => void,
): void {
  useEffect(() => {
    const element = container.current
    if (element === null) return
    const follow = (event: Event): void => {
      const area = (event as CustomEvent<(RangeArea & { type?: string }) | null>).detail
      if (area === null || area.type !== SHEET_ROWS_PANE) return
      const layer = event.composedPath()[0]
      const pane = layer instanceof Element ? (layer.getAttribute('coltype') ?? undefined) : undefined
      const { frozenColumns } = current.current
      setRange(
        normalizeRange({
          fromRow: area.y,
          toRow: area.y1,
          fromColumn: sheetColumnOf(area.x, pane, frozenColumns),
          toColumn: sheetColumnOf(area.x1, pane, frozenColumns),
        }),
      )
    }
    element.addEventListener('setrange', follow)
    return () => element.removeEventListener('setrange', follow)
  }, [container, current, setRange])
}

/**
 * The grid cancels the press on a cell, so the browser would leave the focus in the field typed into
 * last, such as the assistant's request, and Edit > Undo would undo that field; and the text selected
 * last, such as the assistant's answer, which Copy would take instead of the cells. A cell being
 * edited lives inside the grid and keeps the focus.
 */
function useGridTakesFocus(container: { readonly current: HTMLDivElement | null }): void {
  useEffect(() => {
    const element = container.current
    if (element === null) return
    const takeFocus = (event: MouseEvent): void => {
      const grid = event.target instanceof Element ? event.target.closest('revo-grid') : null
      if (grid === null) return
      document.getSelection()?.removeAllRanges()
      const focused = document.activeElement
      if (!(focused instanceof HTMLElement) || grid.contains(focused)) return
      focused.blur()
    }
    element.addEventListener('mousedown', takeFocus, { capture: true })
    return () => element.removeEventListener('mousedown', takeFocus, { capture: true })
  }, [container])
}

function useSheetMenu(
  container: { readonly current: HTMLDivElement | null },
  selection: { readonly current: Range },
  setRange: (range: Range) => void,
): {
  menu: MenuPosition | null
  handleContextMenu: (event: React.MouseEvent<HTMLDivElement>) => void
  closeMenu: () => void
} {
  const [menu, setMenu] = useState<MenuPosition | null>(null)

  // The grid moves the selection on any button's press: inside the selection the right button keeps
  // it, so the menu acts on all of it.
  useEffect(() => {
    const element = container.current
    if (element === null) return
    const keepSelection = (event: MouseEvent): void => {
      if (event.button !== SECONDARY_BUTTON) return
      const position = gridPositionOf(event)
      if (position !== null && rangeContains(selection.current, position.row, position.column)) {
        event.stopPropagation()
      }
    }
    element.addEventListener('mousedown', keepSelection, { capture: true })
    return () => element.removeEventListener('mousedown', keepSelection, { capture: true })
  }, [container, selection])

  /** As in Excel: inside the selection it acts on it; outside, on the clicked cell. */
  const handleContextMenu = useCallback((event: React.MouseEvent<HTMLDivElement>) => {
    event.preventDefault()

    // Headers and empty areas carry no position.
    const position = gridPositionOf(event.nativeEvent)
    if (position !== null && !rangeContains(selection.current, position.row, position.column)) {
      setRange(singleCell(position.row, position.column))
    }

    setMenu({ x: event.clientX, y: event.clientY })
  }, [])

  const closeMenu = useCallback(() => setMenu(null), [])

  return { menu, handleContextMenu, closeMenu }
}

/** Grid, toolbars, menu and shortcuts write: all go through here, so none escapes read-only. */
function useGuardedWrites(
  onChange: (sheet: Sheet) => void,
  onStructure: (change: StructuralChange) => void,
  readOnly: boolean,
): { applyChange: (sheet: Sheet) => void; applyStructure: (change: StructuralChange) => void } {
  const applyChange = useCallback(
    (next: Sheet) => {
      if (!readOnly) onChange(next)
    },
    [onChange, readOnly],
  )

  const applyStructure = useCallback(
    (change: StructuralChange) => {
      if (!readOnly) onStructure(change)
    },
    [onStructure, readOnly],
  )

  return { applyChange, applyStructure }
}
