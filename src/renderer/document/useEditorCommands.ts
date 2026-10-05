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
  /** Bypasses the lock: it is for closing and for the toolbar. */
  readonly setDialog: (dialog: keyof EditorDialogs, open: boolean) => void
  readonly run: (command: EditorCommand) => void
  readonly equationTarget: EquationTarget
}

/**
 * Read-only is checked in `run`, and only there. The `COMMANDS` table is exhaustive: a command
 * without an entry does not compile.
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

      // A cursor just moved may exist only in the DOM, and the menu command arrives before
      // `selectionchange`; a focused note has the same delay.
      if (editor !== null) {
        flushSelection(editor)
        flushNoteSelection(editor.view)
      }

      COMMANDS[command](
        { editor, setDialog, setEquationTarget, referenceContext, pasteWithoutFormat },
        command,
      )
    },
    [editor, readOnly, pasteWithoutFormat, referenceContext, setDialog],
  )

  useEffect(() => onEditorCommand(run), [run])

  return { dialogs, setDialog, run, equationTarget }
}

interface CommandScope {
  readonly editor: Editor | null
  readonly setDialog: (dialog: keyof EditorDialogs, open: boolean) => void
  readonly setEquationTarget: (target: EquationTarget) => void
  readonly referenceContext: () => ReferenceContext
  readonly pasteWithoutFormat: () => Promise<void>
}

type CommandHandler = (scope: CommandScope, command: EditorCommand) => void

const openDialog =
  (dialog: keyof EditorDialogs): CommandHandler =>
  (scope) =>
    scope.setDialog(dialog, true)

const withEditor =
  (action: (editor: Editor, scope: CommandScope, command: EditorCommand) => void): CommandHandler =>
  (scope, command) => {
    if (scope.editor !== null) action(scope.editor, scope, command)
  }

const insertEquation: CommandHandler = (scope, command) => {
  scope.setEquationTarget({ kind: 'insert', display: command === EditorCommand.InsertDisplayEquation })
  scope.setDialog('equation', true)
}

const editEquation: CommandHandler = (scope) => {
  const pos = scope.editor === null ? null : equationAtSelection(scope.editor.state)
  if (pos === null) return
  scope.setEquationTarget({ kind: 'edit', pos })
  scope.setDialog('equation', true)
}

const tableAction = withEditor((editor, _scope, command) => runTableAction(editor, command as TableAction))

const COMMANDS: Readonly<Record<EditorCommand, CommandHandler>> = {
  [EditorCommand.FindReplace]: openDialog('find'),
  [EditorCommand.PageSetup]: openDialog('pageSetup'),
  [EditorCommand.ParagraphSetup]: openDialog('paragraph'),
  [EditorCommand.WordCount]: openDialog('wordCount'),
  [EditorCommand.DocumentProperties]: openDialog('properties'),
  [EditorCommand.SpecialCharacter]: openDialog('specialCharacter'),
  [EditorCommand.InsertEquation]: insertEquation,
  [EditorCommand.InsertDisplayEquation]: insertEquation,
  [EditorCommand.EditEquation]: editEquation,
  [EditorCommand.ImageProperties]: openDialog('imageProperties'),
  [EditorCommand.InsertBookmark]: openDialog('bookmark'),
  [EditorCommand.InsertComment]: withEditor((editor) => insertComment(editor)),
  [EditorCommand.InsertFootnote]: withEditor((editor) => insertNote(editor, NoteKind.Footnote)),
  [EditorCommand.InsertEndnote]: withEditor((editor) => insertNote(editor, NoteKind.Endnote)),
  [EditorCommand.NextComment]: withEditor((editor) => goToComment(editor, 1)),
  [EditorCommand.PreviousComment]: withEditor((editor) => goToComment(editor, -1)),
  [EditorCommand.AuthorName]: openDialog('authorName'),
  [EditorCommand.AcceptChange]: withEditor((editor) => settleChange(editor, true)),
  [EditorCommand.RejectChange]: withEditor((editor) => settleChange(editor, false)),
  [EditorCommand.AcceptAllChanges]: withEditor((editor) => settleAll(editor, true)),
  [EditorCommand.RejectAllChanges]: withEditor((editor) => settleAll(editor, false)),
  [EditorCommand.NextChange]: withEditor((editor) => goToChange(editor, 1)),
  [EditorCommand.PreviousChange]: withEditor((editor) => goToChange(editor, -1)),
  [EditorCommand.ToggleTrackChanges]: () => useWorkspace.getState().toggleTrackChanges(),
  [EditorCommand.InsertTableOfContents]: withEditor((editor, scope) =>
    insertTableOfContents(editor, scope.referenceContext()),
  ),
  [EditorCommand.UpdateTableOfContents]: withEditor((editor, scope) =>
    updateTableOfContents(editor, scope.referenceContext()),
  ),
  [EditorCommand.InsertCaption]: openDialog('caption'),
  [EditorCommand.InsertCrossReference]: openDialog('crossReference'),
  [EditorCommand.UpdateFields]: withEditor((editor, scope) => {
    updateFields(editor, scope.referenceContext())
  }),
  [EditorCommand.PasteWithoutFormat]: (scope) => void scope.pasteWithoutFormat(),
  [EditorCommand.InsertPageBreak]: withEditor((editor) => editor.chain().focus().setPageBreak().run()),
  [EditorCommand.InsertSectionNextPage]: withEditor((editor) => insertSectionBreak(editor, 'nextPage')),
  [EditorCommand.InsertSectionContinuous]: withEditor((editor) => insertSectionBreak(editor, 'continuous')),
  [EditorCommand.InsertSectionEvenPage]: withEditor((editor) => insertSectionBreak(editor, 'evenPage')),
  [EditorCommand.InsertSectionOddPage]: withEditor((editor) => insertSectionBreak(editor, 'oddPage')),
  [EditorCommand.DeleteSectionBreak]: withEditor((editor) => deleteSectionBreak(editor)),
  [EditorCommand.InsertColumnBreak]: withEditor((editor) => insertColumnBreak(editor)),
  [EditorCommand.FormatColumns]: (scope) => {
    if (sectionEditsAllowed()) scope.setDialog('columns', true)
  },

  [TableAction.Insert]: openDialog('table'),
  [TableAction.Properties]: openDialog('tableProperties'),
  [TableAction.RowBefore]: tableAction,
  [TableAction.RowAfter]: tableAction,
  [TableAction.DeleteRow]: tableAction,
  [TableAction.ColumnBefore]: tableAction,
  [TableAction.ColumnAfter]: tableAction,
  [TableAction.DeleteColumn]: tableAction,
  [TableAction.MergeCells]: tableAction,
  [TableAction.SplitCell]: tableAction,
  [TableAction.ToggleHeaderRow]: tableAction,
  [TableAction.Delete]: tableAction,
}
