import { Node, mergeAttributes } from '@tiptap/core'

/**
 * Quebra de página manual.
 *
 * Na tela é uma linha tracejada com rótulo — o editor rola de forma contínua,
 * sem paginar ao vivo (ver docs/00-plano-tecnico.md §6.3). Na exportação da
 * Fase 3 este nó vira `break-after: page` no CSS de impressão, e na Fase 4
 * vira `<w:br w:type="page"/>` no DOCX.
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
  // Bloco indivisível: o cursor o seleciona inteiro, nunca entra "dentro".
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
        // Rótulo por CSS: não entra no texto e não é copiado junto.
        'aria-label': 'Quebra de página',
      }),
    ]
  },

  addCommands() {
    return {
      setPageBreak:
        () =>
        ({ commands }) =>
          commands.insertContent({ type: this.name }),
    }
  },

  addKeyboardShortcuts() {
    return {
      // Mesmo atalho do Word.
      'Mod-Enter': () => this.editor.commands.setPageBreak(),
    }
  },
})
