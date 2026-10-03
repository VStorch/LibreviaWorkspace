import { Node } from '@tiptap/core'

/** O padrão do Word: títulos 1 a 3, com link. */
export const DEFAULT_TOC_INSTRUCTION = ' TOC \\o "1-3" \\h \\z \\u '

/**
 * No arquivo é um campo `TOC`, quase sempre num `w:sdt`, cujo resultado são
 * parágrafos comuns com link para o `_Toc…` e um `PAGEREF`. O leitor tira o campo
 * dos parágrafos (`instr`); `head` conta os parágrafos antes dele, e `sdt` diz se
 * havia controle de conteúdo (`BodyReader.ReadTableOfContents`). As entradas são
 * editáveis, e "Atualizar sumário" as refaz (`references.ts`).
 */
export const TableOfContents = Node.create({
  name: 'tableOfContents',
  group: 'block',
  content: '(paragraph | heading)+',
  defining: true,
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
