import type { Editor } from '@tiptap/react'
import { TableAction } from '@shared/table-actions.js'

/**
 * O que cada ação de tabela faz no editor.
 *
 * Nenhuma delas é nossa: mesclar, dividir, inserir e remover linha ou coluna são
 * comandos do TableKit, que estavam instalados e sem interface nenhuma desde que
 * a tabela passou a ser lida do `.docx`. Aqui só se liga o id do menu ao comando.
 *
 * As duas que abrem diálogo — inserir e propriedades — **não** estão aqui: elas
 * não mexem no documento, e quem as atende é o componente que desenha o diálogo.
 */
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

    // Os dois diálogos. Devolvem `false` para quem chamou saber que ainda há
    // interface a abrir — e não que o comando falhou.
    case TableAction.Insert:
    case TableAction.Properties:
      return false
  }
}
