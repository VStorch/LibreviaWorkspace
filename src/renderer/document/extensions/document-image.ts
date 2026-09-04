import Image from '@tiptap/extension-image'
import { mergeAttributes } from '@tiptap/core'

/**
 * A imagem do documento, com o tamanho que o documento pediu.
 *
 * O `.docx` diz em `wp:extent` de que tamanho a imagem é **na página**, e esse
 * tamanho não precisa ter a proporção do arquivo: quem arrasta um canto sem
 * travar a proporção estica a imagem, e o Word desenha esticado.
 *
 * Só os atributos `width` e `height` não bastam para reproduzir isso. O HTML os
 * trata como uma proporção de reserva — `aspect-ratio: auto <w>/<h>` — e a
 * palavra `auto` diz que, assim que os bytes chegam, a proporção do arquivo
 * ganha. A imagem então voltava ao natural, e no meio do caminho a largura
 * limitada pela coluna arrastava a altura junto.
 *
 * Declarar a proporção sem `auto` fecha as duas pontas: a caixa é reservada
 * antes de a imagem decodificar — o que a paginação precisa, porque ela mede a
 * folha uma vez e a imagem chega depois — e continua valendo depois, inclusive
 * quando `max-width` encolhe a imagem para caber na coluna.
 */
export const DocumentImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      /**
       * A imagem está ancorada ao parágrafo, e não solta no meio da frase.
       *
       * Muda a altura do bloco, e ela conta na paginação: o parágrafo que
       * **ancora** uma imagem ocupa a linha dele além da imagem, e o que a traz
       * no meio da frase ocupa só a imagem. Medido no LibreOffice: 11,55 pt
       * entre duas capturas seguidas num documento de evidências, que é
       * exatamente uma linha de Arial 10 pt.
       */
      anchored: {
        default: null,
        parseHTML: (element: HTMLElement) => (element.hasAttribute('data-anchored') ? true : null),
        renderHTML: (attributes: Record<string, unknown>) =>
          attributes['anchored'] === true ? { 'data-anchored': '' } : {},
      },
    }
  },

  renderHTML({ HTMLAttributes }) {
    const width = Number(HTMLAttributes['width'])
    const height = Number(HTMLAttributes['height'])
    const proportion =
      Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0
        ? { style: `aspect-ratio: ${width} / ${height}` }
        : {}

    return ['img', mergeAttributes(this.options.HTMLAttributes, HTMLAttributes, proportion)]
  },
})
