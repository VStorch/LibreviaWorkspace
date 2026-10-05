import { Node } from '@tiptap/core'
import { fieldKind } from '@services/document/fields.js'

/**
 * The instruction is what the field **is**; the result, what it showed last time. As text, the page
 * number would become a typed number on the first edit. The result only changes with "Update
 * fields" (F9), in `references.ts`.
 */
export const Field = Node.create({
  name: 'field',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      instr: { default: '', parseHTML: (element) => element.getAttribute('data-instr') ?? '' },
      result: { default: '', parseHTML: (element) => element.getAttribute('data-result') ?? '' },
    }
  },

  parseHTML() {
    return [{ tag: 'span[data-field]' }]
  },

  renderHTML({ node }) {
    const instr = String(node.attrs['instr'] ?? '')
    const result = String(node.attrs['result'] ?? '')
    return [
      'span',
      {
        // For CSS: the table of contents `pageref` goes right with the dot leaders.
        'data-field': fieldKind(instr).toLowerCase(),
        'data-instr': instr,
        'data-result': result,
        class: 'field',
      },
      result,
    ]
  },

  renderText({ node }) {
    return String(node.attrs['result'] ?? '')
  },
})
