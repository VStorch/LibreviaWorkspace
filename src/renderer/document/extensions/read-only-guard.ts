import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'

/**
 * `contenteditable="false"` only holds back keyboard and mouse; Tiptap commands do not ask whether
 * the editor is editable. The last gate after `useEditorCommands`: refuses only transactions that
 * **change the document**.
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
