import type { MessageKey } from './i18n/index.js'

/** O rótulo é chave do catálogo: cada consumidor traduz com o `t` que tem. */
export const TableAction = {
  /** Abre o diálogo que pergunta quantas linhas e colunas. */
  Insert: 'table-insert',
  RowBefore: 'table-row-before',
  RowAfter: 'table-row-after',
  DeleteRow: 'table-delete-row',
  ColumnBefore: 'table-column-before',
  ColumnAfter: 'table-column-after',
  DeleteColumn: 'table-delete-column',
  MergeCells: 'table-merge-cells',
  SplitCell: 'table-split-cell',
  /** `w:tblHeader`: a primeira linha se repete no alto de cada página. */
  ToggleHeaderRow: 'table-header-row',
  Delete: 'table-delete',
  /** Abre o diálogo de largura de coluna, bordas e sombreamento. */
  Properties: 'table-properties',
} as const

export type TableAction = (typeof TableAction)[keyof typeof TableAction]

export interface TableActionInfo {
  readonly id: TableAction
  readonly labelKey: MessageKey
  /** Fora de uma tabela, a ação fica apagada no menu e some do menu de contexto. */
  readonly needsTable: boolean
  /** O menu põe um separador quando o grupo muda. */
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
