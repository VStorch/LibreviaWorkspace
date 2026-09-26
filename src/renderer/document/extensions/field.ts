import { Node } from '@tiptap/core'
import { fieldKind } from '@services/document/fields.js'

/**
 * Um campo do Word (`w:fldChar`/`w:instrText`, ou `w:fldSimple`): a instrução e o
 * último resultado calculado.
 *
 * Um nó atômico, e não texto: a instrução (`PAGEREF _Toc123 \h`, `SEQ Figura`,
 * `REF _Ref4 \h`) é o que o campo **é**, e o resultado só o que ele mostrou da
 * última vez. Como texto, o número da página virava número digitado na primeira
 * edição do parágrafo — que era por que o documento com campo abria travado.
 *
 * O resultado é desenhado como está até alguém pedir "Atualizar campos" (F9),
 * como no Word; quem o recalcula é `field-update.ts`. As marcas do nó são a
 * formatação do resultado inteiro, e o link do sumário mora nelas também.
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
        // O tipo do campo em minúsculas, para o CSS: é o `pageref` que o sumário
        // alinha à direita com os pontinhos.
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
