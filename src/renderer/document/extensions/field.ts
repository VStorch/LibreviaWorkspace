import { Node } from '@tiptap/core'
import { fieldKind } from '@services/document/fields.js'

/**
 * A instrução é o que o campo **é**; o resultado, o que mostrou da última vez.
 * Como texto, o número da página viraria número digitado na primeira edição. O
 * resultado só muda com "Atualizar campos" (F9), em `references.ts`.
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
        // Para o CSS: o `pageref` do sumário vai à direita com os pontinhos.
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
