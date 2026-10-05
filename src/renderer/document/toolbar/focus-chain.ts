import type { ChainedCommands, Editor } from '@tiptap/core'

/** Always with `focus()`: otherwise each toolbar click would take the cursor from the editor. */
export function focusChain(editor: Editor): ChainedCommands {
  return editor.chain().focus()
}
