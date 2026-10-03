/**
 * O menu chega ao renderer pelo `App`, que não tem o editor. Os identificadores
 * são os mesmos de `MenuCommand`, para o `App` repassar pelo nome sem tradução.
 */

import { TableAction } from '@shared/table-actions.js'
import type { MenuCommand } from '@shared/types.js'

/** O `satisfies` faz o valor que não existe em `MenuCommand` não compilar. */
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

/** O `App` só sabe de arquivos: seleção, cursor e diálogo do documento são do editor. */
export function asEditorCommand(command: string): EditorCommand | null {
  return KNOWN.has(command) ? (command as EditorCommand) : null
}

/** A lista é a das exceções: um comando novo nasce bloqueado no somente leitura. */
const READS_ONLY: ReadonlySet<EditorCommand> = new Set<EditorCommand>([
  EditorCommand.FindReplace,
  EditorCommand.WordCount,
  // O diálogo abre para "Ir para"; adicionar e excluir se apagam lá dentro.
  EditorCommand.InsertBookmark,
  // Andar entre os comentários só lê; e o nome do autor é preferência, não documento.
  EditorCommand.NextComment,
  EditorCommand.PreviousComment,
  EditorCommand.AuthorName,
  // Andar entre as alterações também só lê; aceitar e rejeitar editam.
  EditorCommand.NextChange,
  EditorCommand.PreviousChange,
  // A equação abre para ser vista; o diálogo só grava com o documento editável.
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
