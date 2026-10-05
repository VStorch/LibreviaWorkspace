import type { MessageKey } from './i18n/index.js'

/** The label is a catalog key: each consumer translates with its own `t`. */
export const TableAction = {
  Insert: 'table-insert',
  RowBefore: 'table-row-before',
  RowAfter: 'table-row-after',
  DeleteRow: 'table-delete-row',
  ColumnBefore: 'table-column-before',
  ColumnAfter: 'table-column-after',
  DeleteColumn: 'table-delete-column',
  MergeCells: 'table-merge-cells',
  SplitCell: 'table-split-cell',
  /** `w:tblHeader`: the first row repeats at the top of each page. */
  ToggleHeaderRow: 'table-header-row',
  Delete: 'table-delete',
  Properties: 'table-properties',
} as const

export type TableAction = (typeof TableAction)[keyof typeof TableAction]

export interface TableActionInfo {
  readonly id: TableAction
  readonly labelKey: MessageKey
  /** Outside a table the action is greyed out in the menu and hidden from the context menu. */
  readonly needsTable: boolean
  /** The menu adds a separator when the group changes. */
  readonly group: number
}

export const TABLE_ACTIONS: readonly TableActionInfo[] = [
  { id: TableAction.Insert, labelKey: 'table.insert', needsTable: false, group: 0 },
  { id: TableAction.RowBefore, labelKey: 'table.rowBefore', needsTable: true, group: 1 },
  { id: TableAction.RowAfter, labelKey: 'table.rowAfter', needsTable: true, group: 1 },
  { id: TableAction.ColumnBefore, labelKey: 'table.columnBefore', needsTable: true, group: 1 },
  { id: TableAction.ColumnAfter, labelKey: 'table.columnAfter', needsTable: true, group: 1 },
  { id: TableAction.DeleteRow, labelKey: 'table.deleteRow', needsTable: true, group: 2 },
  { id: TableAction.DeleteColumn, labelKey: 'table.deleteColumn', needsTable: true, group: 2 },
  { id: TableAction.Delete, labelKey: 'table.delete', needsTable: true, group: 2 },
  { id: TableAction.MergeCells, labelKey: 'table.mergeCells', needsTable: true, group: 3 },
  { id: TableAction.SplitCell, labelKey: 'table.splitCell', needsTable: true, group: 3 },
  { id: TableAction.ToggleHeaderRow, labelKey: 'table.headerRow', needsTable: true, group: 4 },
  { id: TableAction.Properties, labelKey: 'table.properties', needsTable: true, group: 4 },
]
