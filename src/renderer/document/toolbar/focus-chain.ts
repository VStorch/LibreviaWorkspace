import type { ChainedCommands, Editor } from '@tiptap/core'

/**
 * O encadeamento que todo botão da barra usa.
 *
 * Sempre com `focus()`: o comando roda sobre a seleção do editor, e sem devolver o
 * foco a ele o usuário perde o cursor a cada clique na barra.
 */
export function focusChain(editor: Editor): ChainedCommands {
  return editor.chain().focus()
}
