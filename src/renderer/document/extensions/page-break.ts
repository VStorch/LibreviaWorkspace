import { Node, mergeAttributes } from '@tiptap/core'
import { t } from '../../i18n.js'

/**
 * Não desenha nada: a folha termina ali, como na vista de impressão do Word e do
 * LibreOffice. Continua selecionável e apagável com Backspace no começo da folha
 * seguinte. No DOCX vira `<w:br w:type="page"/>`.
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
          // Com o parágrafo seguinte e o cursor dentro dele, como o Ctrl+Enter do
          // Word: o nó atômico selecionado seria substituído pela primeira tecla.
          commands.insertContent([{ type: this.name }, { type: 'paragraph' }]),
    }
  },

  addKeyboardShortcuts() {
    return {
      'Mod-Enter': () => this.editor.commands.setPageBreak(),
    }
  },
})
