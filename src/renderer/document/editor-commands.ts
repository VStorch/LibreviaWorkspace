/**
 * The menu reaches the renderer through `App`, which has no editor. The identifiers match
 * `MenuCommand`, so `App` forwards them by name without translation.
 */

import { TableAction } from '@shared/table-actions.js'
import type { MenuCommand } from '@shared/types.js'

/** `satisfies` makes a value missing from `MenuCommand` fail to compile. */
export const EditorCommand = {
  FindReplace: 'find-replace',
  PageSetup: 'page-setup',
  InsertPageBreak: 'insert-page-break',
  InsertSectionNextPage: 'insert-section-next-page',
  InsertSectionContinuous: 'insert-section-continuous',
  InsertSectionEvenPage: 'insert-section-even-page',
  InsertSectionOddPage: 'insert-section-odd-page',
  DeleteSectionBreak: 'delete-section-break',
  InsertColumnBreak: 'insert-column-break',
  FormatColumns: 'format-columns',
  ParagraphSetup: 'paragraph-setup',
  PasteWithoutFormat: 'paste-without-format',
  WordCount: 'word-count',
  DocumentProperties: 'document-properties',
  SpecialCharacter: 'special-character',
  InsertEquation: 'insert-equation',
  InsertDisplayEquation: 'insert-display-equation',
  EditEquation: 'edit-equation',
  ImageProperties: 'image-properties',
  InsertBookmark: 'insert-bookmark',
  InsertTableOfContents: 'insert-table-of-contents',
  UpdateTableOfContents: 'update-table-of-contents',
  UpdateFields: 'update-fields',
  InsertCaption: 'insert-caption',
  InsertCrossReference: 'insert-cross-reference',
  InsertComment: 'insert-comment',
  InsertFootnote: 'insert-footnote',
  InsertEndnote: 'insert-endnote',
  NextComment: 'next-comment',
  PreviousComment: 'previous-comment',
  AuthorName: 'author-name',
  AcceptChange: 'accept-change',
  RejectChange: 'reject-change',
  AcceptAllChanges: 'accept-all-changes',
  RejectAllChanges: 'reject-all-changes',
  NextChange: 'next-change',
  PreviousChange: 'previous-change',
  ToggleTrackChanges: 'toggle-track-changes',
  ...TableAction,
} as const satisfies Record<string, MenuCommand>

export type EditorCommand = (typeof EditorCommand)[keyof typeof EditorCommand]

const KNOWN = new Set<string>(Object.values(EditorCommand))

/** `App` only knows about files: selection, cursor and document dialogs belong to the editor. */
export function asEditorCommand(command: string): EditorCommand | null {
  return KNOWN.has(command) ? (command as EditorCommand) : null
}

/** The list holds the exceptions: a new command starts blocked when read-only. */
const READS_ONLY: ReadonlySet<EditorCommand> = new Set<EditorCommand>([
  EditorCommand.FindReplace,
  EditorCommand.WordCount,
  // The dialog opens for "Go to"; add and delete grey out inside it.
  EditorCommand.InsertBookmark,
  // Moving between comments only reads; the author name is a preference, not the document.
  EditorCommand.NextComment,
  EditorCommand.PreviousComment,
  EditorCommand.AuthorName,
  // Moving between changes also only reads; accepting and rejecting edit.
  EditorCommand.NextChange,
  EditorCommand.PreviousChange,
  // The equation opens for viewing; the dialog only saves when the document is editable.
  EditorCommand.EditEquation,
])

export function runsWhileLocked(command: EditorCommand): boolean {
  return READS_ONLY.has(command)
}

const listeners = new Set<(command: EditorCommand) => void>()

export function onEditorCommand(listener: (command: EditorCommand) => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}

export function emitEditorCommand(command: EditorCommand): void {
  for (const listener of listeners) listener(command)
}
