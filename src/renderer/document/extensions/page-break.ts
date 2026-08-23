import { Node, mergeAttributes } from '@tiptap/core'

/**
 * Quebra de página manual.
 *
 * O editor pagina ao vivo, então a quebra é visível pelo que ela faz: a folha
 * termina ali, e é só isso que o Word e o LibreOffice mostram na vista de
 * impressão. A linha tracejada escrita "QUEBRA DE PÁGINA" fazia sentido quando
 * a tela era uma tira contínua e a quebra não tinha efeito nenhum de se ver;
 * hoje ela é ruído, e na capa do modelo de manual caía no meio do desenho.
 *
 * O nó continua existindo, selecionável e apagável — com Backspace no começo da
 * folha seguinte, como no Word.
 *
 * Na exportação vira `break-after: page` no CSS de impressão, e no DOCX vira
 * `<w:br w:type="page"/>`.
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
          // A quebra vem acompanhada do parágrafo que a sucede, e o cursor
          // termina dentro dele. É o que o Word faz com Ctrl+Enter, e resolve um
          // defeito que só ficou visível quando a tela passou a paginar:
          // `insertContent` sozinho deixa o nó **selecionado**, e como ele é
          // atômico a primeira tecla digitada o substitui — quem inseria a
          // quebra e continuava escrevendo a apagava sem sinal nenhum.
          commands.insertContent([{ type: this.name }, { type: 'paragraph' }]),
    }
  },

  addKeyboardShortcuts() {
    return {
      // Mesmo atalho do Word.
      'Mod-Enter': () => this.editor.commands.setPageBreak(),
    }
  },
})
