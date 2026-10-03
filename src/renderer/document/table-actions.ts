import type { Editor } from '@tiptap/react'
import { TableAction } from '@shared/table-actions.js'

/** Comandos do TableKit; inserir e propriedades abrem diálogo e não estão aqui. */
export function runTableAction(editor: Editor, action: TableAction): boolean {
  const chain = editor.chain().focus()

  switch (action) {
    case TableAction.RowBefore:
      return chain.addRowBefore().run()
    case TableAction.RowAfter:
      return chain.addRowAfter().run()
    case TableAction.DeleteRow:
      return chain.deleteRow().run()
    case TableAction.ColumnBefore:
      return chain.addColumnBefore().run()
    case TableAction.ColumnAfter:
      return chain.addColumnAfter().run()
    case TableAction.DeleteColumn:
      return chain.deleteColumn().run()
    case TableAction.MergeCells:
      return chain.mergeCells().run()
    case TableAction.SplitCell:
      return chain.splitCell().run()
    case TableAction.ToggleHeaderRow:
      return chain.toggleHeaderRow().run()
    case TableAction.Delete:
      return chain.deleteTable().run()

    // `false`: ainda há interface a abrir, e não falha.
    case TableAction.Insert:
    case TableAction.Properties:
      return false
  }
}
