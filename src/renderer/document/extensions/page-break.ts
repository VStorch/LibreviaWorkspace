import { Node, mergeAttributes } from '@tiptap/core'
import { t } from '../../i18n.js'

/**
 * Draws nothing: the sheet ends there, as in Word's and LibreOffice's print view. Still selectable
 * and deletable with Backspace at the start of the next sheet. In DOCX it becomes `<w:br
 * w:type="page"/>`.
 */
declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    pageBreak: {
      setPageBreak: () => ReturnType
    }
  }
}

export const PageBreak = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  selectable: true,

  parseHTML() {
    return [{ tag: 'div[data-page-break]' }]
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'div',
      mergeAttributes(HTMLAttributes, {
        'data-page-break': '',
        class: 'page-break',
        'aria-label': t('menu.insert.pageBreak'),
      }),
    ]
  },

  addCommands() {
    return {
      setPageBreak:
        () =>
        ({ commands }) =>
          // With the following paragraph and the cursor in it, like Word's Ctrl+Enter: a selected
          // atomic node would be replaced by the first key press.
          commands.insertContent([{ type: this.name }, { type: 'paragraph' }]),
    }
  },

  addKeyboardShortcuts() {
    return {
      'Mod-Enter': () => this.editor.commands.setPageBreak(),
    }
  },
})
