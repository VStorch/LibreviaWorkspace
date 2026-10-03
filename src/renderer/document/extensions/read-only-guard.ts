import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'

/**
 * O `contenteditable="false"` só segura teclado e mouse; os comandos do Tiptap
 * não perguntam se o editor é editável. Último portão depois de
 * `useEditorCommands`: recusa só a transação que **muda o documento**.
 */
export const ReadOnlyGuard = Extension.create({
  name: 'readOnlyGuard',

  addProseMirrorPlugins() {
    const { editor } = this
    return [
      new Plugin({
        key: new PluginKey('readOnlyGuard'),
        filterTransaction: (transaction) => !transaction.docChanged || editor.isEditable,
      }),
    ]
  },
})
