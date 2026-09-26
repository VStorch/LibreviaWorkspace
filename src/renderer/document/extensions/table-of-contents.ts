import { Node } from '@tiptap/core'

/** A instrução do sumário que o Word insere por padrão: títulos 1 a 3, com link. */
export const DEFAULT_TOC_INSTRUCTION = ' TOC \\o "1-3" \\h \\z \\u '

/**
 * O sumário: um bloco com os parágrafos das entradas dentro.
 *
 * No arquivo ele é um campo `TOC` — quase sempre dentro de um controle de
 * conteúdo (`w:sdt`) — cujo resultado são parágrafos comuns: o título, e uma
 * entrada por título do documento, cada uma com o link para o marcador `_Toc…`
 * do título e um `PAGEREF` com o número da página. O leitor tira o campo dos
 * parágrafos e o põe aqui (`instr`); `head` conta os parágrafos antes dele, e
 * `sdt` diz se havia controle de conteúdo em volta. Ver
 * `BodyReader.ReadTableOfContents`.
 *
 * As entradas continuam editáveis, como no Word, e "Atualizar sumário" as refaz
 * dos títulos — ver `toc-update.ts`.
 */
export const TableOfContents = Node.create({
  name: 'tableOfContents',
  group: 'block',
  content: '(paragraph | heading)+',
  defining: true,
  // Apagar até a borda não funde o sumário com o parágrafo vizinho.
  isolating: true,

  addAttributes() {
    return {
      instr: { default: DEFAULT_TOC_INSTRUCTION, parseHTML: (element) => element.getAttribute('data-instr') },
      head: {
        default: 0,
        parseHTML: (element) => Number(element.getAttribute('data-head') ?? 0),
      },
      sdt: { default: true, parseHTML: (element) => element.getAttribute('data-sdt') !== 'false' },
    }
  },

  parseHTML() {
    return [{ tag: 'div[data-toc]' }]
  },

  renderHTML({ node }) {
    return [
      'div',
      {
        'data-toc': '',
        'data-instr': String(node.attrs['instr']),
        'data-head': String(node.attrs['head']),
        'data-sdt': String(node.attrs['sdt']),
        class: 'toc',
      },
      0,
    ]
  },
})
