import { Node } from '@tiptap/core'

/** Word's default: headings 1 to 3, with links. */
export const DEFAULT_TOC_INSTRUCTION = ' TOC \\o "1-3" \\h \\z \\u '

/**
 * In the file it is a `TOC` field, almost always in a `w:sdt`, whose result is plain paragraphs
 * linking to the `_Toc…` and a `PAGEREF`. The reader takes the field from the paragraphs (`instr`);
 * `head` counts the paragraphs before it, and `sdt` says whether there was a content control
 * (`BodyReader.ReadTableOfContents`). Entries are editable, and "Update table" rebuilds them
 * (`references.ts`).
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
