import Image from '@tiptap/extension-image'
import { mergeAttributes } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { ReactNodeViewRenderer } from '@tiptap/react'
import { ImageNodeView } from '../ImageNodeView.js'

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
 *
 * ## O `NodeView`
 *
 * As alças de redimensionamento pedem elementos ao lado da imagem, e um `<img>`
 * sozinho não tem onde os pôr. O `renderHTML` abaixo continua existindo e é o que
 * vale no HTML da impressão e do PDF: lá não há alça nenhuma a desenhar.
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

      /**
       * Alinhamento da imagem **em bloco**, na coluna de texto.
       *
       * Só vale para a imagem sem parágrafo, que o schema de hoje não produz: a
       * imagem é conteúdo de linha, e quem a alinha é o `w:jc` do parágrafo dela
       * — no OOXML não existe imagem centralizada, existe parágrafo centralizado
       * com uma imagem dentro. Fica para ler o que foi salvo quando a imagem da
       * barra ainda era um bloco, e o sidecar ainda sabe gravá-la.
       */
      align: {
        default: null,
        parseHTML: (element: HTMLElement) => element.getAttribute('data-align'),
        renderHTML: (attributes: Record<string, unknown>) => {
          const align = attributes['align']
          if (typeof align !== 'string' || align === '') return {}
          return { 'data-align': align }
        },
      },
    }
  },

  addNodeView() {
    return ReactNodeViewRenderer(ImageNodeView, {
      // A marca de ancorada vai no elemento de fora, que é o filho do parágrafo.
      // As regras da folha de estilo que tiram a imagem ancorada da linha e dão a
      // linha dela ao parágrafo olham para o filho direto do `p`, e a moldura das
      // alças mora um nível abaixo — nela a marca não alcançava regra nenhuma.
      // Estático de propósito: o `ReactRenderer` só acrescenta atributos, e ser
      // ancorada é coisa que só o arquivo decide.
      attrs: ({ node }) => (node.attrs['anchored'] === true ? { 'data-anchored': '' } : {}),
    })
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

/** A imagem em que o cursor está, com a posição dela no documento. */
export interface PlacedImage {
  readonly node: ProseMirrorNode
  readonly pos: number
  /** O parágrafo que a abriga, quando há um: é ele que alinha a imagem lida do arquivo. */
  readonly paragraphPos: number | null
}

/**
 * A imagem selecionada, ou a que estiver dentro do bloco do cursor.
 *
 * Os dois casos existem: clicar na imagem faz uma seleção de nó, e chegar nela
 * com as setas deixa o cursor no parágrafo que a contém. Sem o segundo caso o
 * item de menu ficava apagado justamente para quem navega pelo teclado.
 */
export function imageAt(editor: Editor): PlacedImage | null {
  const { selection } = editor.state
  const selected = (selection as { node?: ProseMirrorNode }).node

  if (selected?.type.name === 'image') {
    const $from = selection.$from
    const parent = $from.depth > 0 ? $from.node($from.depth) : null
    return {
      node: selected,
      pos: selection.from,
      paragraphPos: parent !== null && isAligned(parent) ? $from.before($from.depth) : null,
    }
  }

  const $from = selection.$from
  const parent = $from.parent
  let found: PlacedImage | null = null

  parent.forEach((child, offset) => {
    if (found !== null || child.type.name !== 'image') return
    found = {
      node: child,
      pos: $from.start() + offset,
      paragraphPos: isAligned(parent) ? $from.before($from.depth) : null,
    }
  })

  return found
}

/** O bloco que tem alinhamento próprio — os tipos que o TextAlign conhece. */
function isAligned(block: ProseMirrorNode): boolean {
  return block.type.name === 'paragraph' || block.type.name === 'heading'
}

/**
 * Texto alternativo e alinhamento da imagem, numa transação.
 *
 * O alinhamento vai para o **parágrafo** quando há um, e para o atributo da
 * imagem quando ela é um bloco solto. São os dois caminhos que o arquivo tem, e
 * escolher o errado gravaria um alinhamento que ninguém veria de volta.
 */
export function applyImageProperties(
  editor: Editor,
  placed: PlacedImage,
  properties: { readonly alt: string; readonly align: string | null },
): boolean {
  const chain = editor.chain().focus()

  chain.command(({ tr }) => {
    const node = tr.doc.nodeAt(placed.pos)
    if (node === null || node.type.name !== 'image') return false

    // Passo de atributo, e não `setNodeMarkup`: numa folha este substitui o nó,
    // e a seleção da imagem não sobrevive — ver ImageNodeView.
    tr.setNodeAttribute(placed.pos, 'alt', properties.alt === '' ? null : properties.alt)
    tr.setNodeAttribute(placed.pos, 'align', placed.paragraphPos === null ? properties.align : null)
    return true
  })

  if (placed.paragraphPos !== null) {
    const { paragraphPos } = placed
    chain.command(({ tr }) => {
      if (tr.doc.nodeAt(paragraphPos) === null) return false
      tr.setNodeAttribute(paragraphPos, 'textAlign', properties.align)
      return true
    })
  }

  return chain.run()
}
