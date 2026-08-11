/**
 * Canal interno entre o menu nativo e o editor.
 *
 * O menu vive no processo main e chega ao renderer pelo `App`, que não tem
 * referência ao editor. Um emissor mínimo resolve isso sem colocar o objeto do
 * editor — que não é serializável — dentro do store.
 */

export type EditorCommand = 'find-replace' | 'page-setup' | 'insert-page-break'

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
