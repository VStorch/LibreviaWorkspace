import type { ChainedCommands, Editor } from '@tiptap/core'

/** Sempre com `focus()`: senão cada clique na barra tiraria o cursor do editor. */
export function focusChain(editor: Editor): ChainedCommands {
  return editor.chain().focus()
}
