import { useCallback, useMemo, useRef, useState } from 'react'
import { RevoGrid, type ColumnRegular } from '@revolist/react-datagrid'
import type {
  AfterEditEvent,
  BeforeSaveDataDetails,
  ChangedRange,
  FocusAfterRenderEvent,
  RevoGridCustomEvent,
} from '@revolist/revogrid'
import { formatCell } from '@services/spreadsheet/format.js'
import { DEFAULT_COLUMN_WIDTH, columnName, getCell, type Sheet } from '@services/spreadsheet/model.js'
import {
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
import { gridPositionOf } from './grid-position.js'
import { useFormatShortcuts } from './useFormatShortcuts.js'
import { useTypeAhead } from './useTypeAhead.js'

/**
 * The grid works with object rows, and the model is a sparse map by A1 reference: the translation
 * lives here.
 */

type GridRow = Record<string, string>

export function SpreadsheetEditor({
  sheet,
  onChange,
  onStructure,
  readOnly = false,
}: {
  sheet: Sheet
  onChange: (sheet: Sheet) => void
  /** A **workbook** operation: a row inserted here changes `=Dados!A5` written on another tab. */
  onStructure: (change: StructuralChange) => void
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

  const typeAhead = useTypeAhead(readOnly)

  const { applyChange, applyStructure } = useGuardedWrites(onChange, onStructure, readOnly)

  const { columns, source } = useGridModel(sheet, current)
  const { handleEdit, handleEditStart } = useEditHandlers(current, applyChange, typeAhead)
  const { handleResize, handleFocus, handleRange, handleAutofill } = useSelectionHandlers(
    current,
    applyChange,
    typeAhead,
    setRange,
  )
  const { menu, handleContextMenu, closeMenu } = useSheetMenu(selection, setRange)

  useFormatShortcuts((key) => applyChange(toggleStyle(current.current, selection.current, key)))

  return (
    <div className="sheet" onContextMenu={handleContextMenu}>
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
        onBeforerange={handleRange}
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

  /** Shows the **formula**: leaving the cell by accident would write the result over it. */
  const handleEditStart = useCallback((event: RevoGridCustomEvent<BeforeSaveDataDetails>) => {
    const detail = event.detail
    const column = Number.parseInt(String(detail.prop).slice(1), 10)
    if (!Number.isInteger(column)) return

    const formula = getCell(current.current, detail.rowIndex, column)?.formula
    if (formula !== undefined) detail.val = formula
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
  handleRange: (event: RevoGridCustomEvent<ChangedRange>) => void
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
      const { rowIndex, colIndex } = event.detail
      if (!Number.isInteger(rowIndex) || !Number.isInteger(colIndex)) return

      setRange(singleCell(rowIndex, colIndex))
      typeAhead.settle()
    },
    [typeAhead],
  )

  const handleRange = useCallback((event: RevoGridCustomEvent<ChangedRange>) => {
    const area = event.detail.newRange
    if (area === null) return
    setRange(normalizeRange({ fromRow: area.y, fromColumn: area.x, toRow: area.y1, toColumn: area.x1 }))
  }, [])

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

  return { handleResize, handleFocus, handleRange, handleAutofill }
}

function useSheetMenu(
  selection: { readonly current: Range },
  setRange: (range: Range) => void,
): {
  menu: MenuPosition | null
  handleContextMenu: (event: React.MouseEvent<HTMLDivElement>) => void
  closeMenu: () => void
} {
  const [menu, setMenu] = useState<MenuPosition | null>(null)

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
