import { COMMANDS } from './commands/registry.js'
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

function action(id: TableAction, needsTable: boolean, group: number): TableActionInfo {
  return { id, labelKey: COMMANDS[id].labelKey, needsTable, group }
}

export const TABLE_ACTIONS: readonly TableActionInfo[] = [
  action(TableAction.Insert, false, 0),
  action(TableAction.RowBefore, true, 1),
  action(TableAction.RowAfter, true, 1),
  action(TableAction.ColumnBefore, true, 1),
  action(TableAction.ColumnAfter, true, 1),
  action(TableAction.DeleteRow, true, 2),
  action(TableAction.DeleteColumn, true, 2),
  action(TableAction.Delete, true, 2),
  action(TableAction.MergeCells, true, 3),
  action(TableAction.SplitCell, true, 3),
  action(TableAction.ToggleHeaderRow, true, 4),
  action(TableAction.Properties, true, 4),
]
