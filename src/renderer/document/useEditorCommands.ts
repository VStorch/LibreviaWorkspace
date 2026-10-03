import { useCallback, useEffect, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { TableAction } from '@shared/table-actions.js'
import { EditorCommand, onEditorCommand, runsWhileLocked } from './editor-commands.js'
import { runTableAction } from './table-actions.js'
import { goToComment, insertComment } from './comment-commands.js'
import { insertNote } from './note-commands.js'
import { flushNoteSelection } from './extensions/note-view.js'
import { NoteKind } from '@services/document/notes.js'
import { goToChange, settleAll, settleChange } from './revision-commands.js'
import { useWorkspace } from '../state/workspace.js'
import { equationAtSelection, type EquationTarget } from './math-commands.js'
import {
  deleteSectionBreak,
  insertColumnBreak,
  insertSectionBreak,
  sectionEditsAllowed,
} from './section-commands.js'
import {
  flushSelection,
  insertTableOfContents,
  updateFields,
  updateTableOfContents,
  type ReferenceContext,
} from './references.js'

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
  readonly listFormat: boolean
  readonly listStart: boolean
  readonly bookmark: boolean
  readonly caption: boolean
  readonly crossReference: boolean
  readonly columns: boolean
  readonly authorName: boolean
  readonly properties: boolean
  readonly equation: boolean
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
  listFormat: false,
  listStart: false,
  bookmark: false,
  caption: false,
  crossReference: false,
  columns: false,
  authorName: false,
  properties: false,
  equation: false,
}

export interface EditorCommands {
  readonly dialogs: EditorDialogs
  /** Sem passar pela trava: é para fechar e para a barra. */
  readonly setDialog: (dialog: keyof EditorDialogs, open: boolean) => void
  readonly run: (command: EditorCommand) => void
  readonly equationTarget: EquationTarget
}

/**
 * O somente leitura é conferido no `run`, e só nele. O `switch` é exaustivo: um
 * comando sem caso não compila.
 */
export function useEditorCommands(
  editor: Editor | null,
  readOnly: boolean,
  pasteWithoutFormat: () => Promise<void>,
  referenceContext: () => ReferenceContext,
): EditorCommands {
  const [dialogs, setDialogs] = useState<EditorDialogs>(CLOSED)
  const [equationTarget, setEquationTarget] = useState<EquationTarget>({ kind: 'insert', display: false })

  const setDialog = useCallback((dialog: keyof EditorDialogs, open: boolean) => {
    setDialogs((current) => (current[dialog] === open ? current : { ...current, [dialog]: open }))
  }, [])

  const run = useCallback(
    (command: EditorCommand) => {
      if (readOnly && !runsWhileLocked(command)) return

      // O cursor recém-movido pode estar só no DOM, e o comando do menu chega
      // antes do `selectionchange`; a nota com foco tem a mesma demora.
      if (editor !== null) {
        flushSelection(editor)
        flushNoteSelection(editor.view)
      }

      switch (command) {
        case EditorCommand.FindReplace:
          return setDialog('find', true)
        case EditorCommand.PageSetup:
          return setDialog('pageSetup', true)
        case EditorCommand.ParagraphSetup:
          return setDialog('paragraph', true)
        case EditorCommand.WordCount:
          return setDialog('wordCount', true)
        case EditorCommand.DocumentProperties:
          return setDialog('properties', true)
        case EditorCommand.SpecialCharacter:
          return setDialog('specialCharacter', true)
        case EditorCommand.InsertEquation:
        case EditorCommand.InsertDisplayEquation:
          setEquationTarget({ kind: 'insert', display: command === EditorCommand.InsertDisplayEquation })
          return setDialog('equation', true)
        case EditorCommand.EditEquation: {
          const pos = editor === null ? null : equationAtSelection(editor.state)
          if (pos === null) return
          setEquationTarget({ kind: 'edit', pos })
          return setDialog('equation', true)
        }
        case EditorCommand.ImageProperties:
          return setDialog('imageProperties', true)
        case EditorCommand.InsertBookmark:
          return setDialog('bookmark', true)
        case EditorCommand.InsertComment:
          if (editor !== null) insertComment(editor)
          return
        case EditorCommand.InsertFootnote:
        case EditorCommand.InsertEndnote:
          if (editor !== null)
            insertNote(
              editor,
              command === EditorCommand.InsertFootnote ? NoteKind.Footnote : NoteKind.Endnote,
            )
          return
        case EditorCommand.NextComment:
        case EditorCommand.PreviousComment:
          if (editor !== null) goToComment(editor, command === EditorCommand.NextComment ? 1 : -1)
          return
        case EditorCommand.AuthorName:
          return setDialog('authorName', true)
        case EditorCommand.AcceptChange:
        case EditorCommand.RejectChange:
          if (editor !== null) settleChange(editor, command === EditorCommand.AcceptChange)
          return
        case EditorCommand.AcceptAllChanges:
        case EditorCommand.RejectAllChanges:
          if (editor !== null) settleAll(editor, command === EditorCommand.AcceptAllChanges)
          return
        case EditorCommand.NextChange:
        case EditorCommand.PreviousChange:
          if (editor !== null) goToChange(editor, command === EditorCommand.NextChange ? 1 : -1)
          return
        case EditorCommand.ToggleTrackChanges:
          return useWorkspace.getState().toggleTrackChanges()
        case EditorCommand.InsertTableOfContents:
          if (editor !== null) insertTableOfContents(editor, referenceContext())
          return
        case EditorCommand.UpdateTableOfContents:
          if (editor !== null) updateTableOfContents(editor, referenceContext())
          return
        case EditorCommand.InsertCaption:
          return setDialog('caption', true)
        case EditorCommand.InsertCrossReference:
          return setDialog('crossReference', true)
        case EditorCommand.UpdateFields:
          if (editor !== null) updateFields(editor, referenceContext())
          return
        case EditorCommand.PasteWithoutFormat:
          void pasteWithoutFormat()
          return
        case EditorCommand.InsertPageBreak:
          editor?.chain().focus().setPageBreak().run()
          return
        case EditorCommand.InsertSectionNextPage:
          if (editor !== null) insertSectionBreak(editor, 'nextPage')
          return
        case EditorCommand.InsertSectionContinuous:
          if (editor !== null) insertSectionBreak(editor, 'continuous')
          return
        case EditorCommand.InsertSectionEvenPage:
          if (editor !== null) insertSectionBreak(editor, 'evenPage')
          return
        case EditorCommand.InsertSectionOddPage:
          if (editor !== null) insertSectionBreak(editor, 'oddPage')
          return
        case EditorCommand.DeleteSectionBreak:
          if (editor !== null) deleteSectionBreak(editor)
          return
        case EditorCommand.InsertColumnBreak:
          if (editor !== null) insertColumnBreak(editor)
          return
        case EditorCommand.FormatColumns:
          if (sectionEditsAllowed()) setDialog('columns', true)
          return

        case TableAction.Insert:
          return setDialog('table', true)
        case TableAction.Properties:
          return setDialog('tableProperties', true)

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
    [editor, readOnly, pasteWithoutFormat, referenceContext, setDialog],
  )

  useEffect(() => onEditorCommand(run), [run])

  return { dialogs, setDialog, run, equationTarget }
}
