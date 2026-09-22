import { useCallback, useEffect, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { TableAction } from '@shared/table-actions.js'
import { EditorCommand, onEditorCommand, runsWhileLocked } from './editor-commands.js'
import { runTableAction } from './table-actions.js'

/** Os diálogos e painéis que um comando do editor abre. */
export interface EditorDialogs {
  readonly find: boolean
  readonly pageSetup: boolean
  readonly paragraph: boolean
  readonly styles: boolean
  readonly wordCount: boolean
  readonly specialCharacter: boolean
  readonly table: boolean
  readonly tableProperties: boolean
  readonly imageProperties: boolean
}

const CLOSED: EditorDialogs = {
  find: false,
  pageSetup: false,
  paragraph: false,
  styles: false,
  wordCount: false,
  specialCharacter: false,
  table: false,
  tableProperties: false,
  imageProperties: false,
}

export interface EditorCommands {
  readonly dialogs: EditorDialogs
  /** Abre ou fecha um diálogo sem passar pela trava — é para fechar e para a barra. */
  readonly setDialog: (dialog: keyof EditorDialogs, open: boolean) => void
  /** Roda um comando do editor, venha do menu nativo ou do botão direito. */
  readonly run: (command: EditorCommand) => void
}

/**
 * Os comandos do editor, e a trava do somente leitura, num lugar só.
 *
 * O menu nativo e o botão direito chegam aqui pelo mesmo `run`, e é nele — e só
 * nele — que o somente leitura é conferido. Antes cada comando decidia por si:
 * colar sem formatação conferia, a quebra de página e as ações de tabela não, e
 * um documento aberto travado ganhava linha, perdia a tabela e virava "Não
 * salvo" pelo menu.
 *
 * O `switch` é exaustivo de propósito: um comando novo em `EditorCommand` que
 * não tenha caso aqui não compila, em vez de chegar e não fazer nada.
 */
export function useEditorCommands(
  editor: Editor | null,
  readOnly: boolean,
  pasteWithoutFormat: () => Promise<void>,
): EditorCommands {
  const [dialogs, setDialogs] = useState<EditorDialogs>(CLOSED)

  const setDialog = useCallback((dialog: keyof EditorDialogs, open: boolean) => {
    setDialogs((current) => (current[dialog] === open ? current : { ...current, [dialog]: open }))
  }, [])

  const run = useCallback(
    (command: EditorCommand) => {
      if (readOnly && !runsWhileLocked(command)) return

      switch (command) {
        case EditorCommand.FindReplace:
          return setDialog('find', true)
        case EditorCommand.PageSetup:
          return setDialog('pageSetup', true)
        case EditorCommand.ParagraphSetup:
          return setDialog('paragraph', true)
        case EditorCommand.WordCount:
          return setDialog('wordCount', true)
        case EditorCommand.SpecialCharacter:
          return setDialog('specialCharacter', true)
        case EditorCommand.ImageProperties:
          return setDialog('imageProperties', true)
        case EditorCommand.PasteWithoutFormat:
          void pasteWithoutFormat()
          return
        case EditorCommand.InsertPageBreak:
          editor?.chain().focus().setPageBreak().run()
          return

        case TableAction.Insert:
          return setDialog('table', true)
        case TableAction.Properties:
          return setDialog('tableProperties', true)

        // O resto das ações de tabela é comando do TableKit, e não diálogo.
        case TableAction.RowBefore:
        case TableAction.RowAfter:
        case TableAction.DeleteRow:
        case TableAction.ColumnBefore:
        case TableAction.ColumnAfter:
        case TableAction.DeleteColumn:
        case TableAction.MergeCells:
        case TableAction.SplitCell:
        case TableAction.ToggleHeaderRow:
        case TableAction.Delete:
          if (editor !== null) runTableAction(editor, command)
          return

        default:
          command satisfies never
      }
    },
    [editor, readOnly, pasteWithoutFormat, setDialog],
  )

  useEffect(() => onEditorCommand(run), [run])

  return { dialogs, setDialog, run }
}
