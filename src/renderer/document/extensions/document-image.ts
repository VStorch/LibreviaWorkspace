import Image from '@tiptap/extension-image'
import { mergeAttributes } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { ReactNodeViewRenderer } from '@tiptap/react'
import { ImageNodeView } from '../ImageNodeView.js'

/**
 * `wp:extent` não precisa ter a proporção do arquivo: o Word desenha esticado.
 * `width` e `height` sozinhos virariam `aspect-ratio: auto`, e a proporção do
 * arquivo ganharia; declarada sem `auto`, a caixa é reservada antes de a imagem
 * decodificar, que é quando a paginação mede. O `NodeView` dá lugar às alças; o
 * `renderHTML` vale na impressão.
 */
export const DocumentImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      /** O parágrafo que ancora a imagem ocupa a linha dele além dela, como mede o LibreOffice. */
      anchored: {
        default: null,
        parseHTML: (element: HTMLElement) => (element.hasAttribute('data-anchored') ? true : null),
        renderHTML: (attributes: Record<string, unknown>) =>
          attributes['anchored'] === true ? { 'data-anchored': '' } : {},
      },

      /** Só da imagem em bloco dos rascunhos antigos: no OOXML quem alinha é o `w:jc` do parágrafo. */
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
      // No elemento de fora, que é o filho do parágrafo: é ele que as regras da folha de estilo olham.
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

export interface PlacedImage {
  readonly node: ProseMirrorNode
  readonly pos: number
  /** É ele que alinha a imagem lida do arquivo. */
  readonly paragraphPos: number | null
}

/** Selecionada, ou no bloco do cursor: chegar pelas setas deixa o cursor no parágrafo. */
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

function isAligned(block: ProseMirrorNode): boolean {
  return block.type.name === 'paragraph' || block.type.name === 'heading'
}

/** O alinhamento vai para o **parágrafo** quando há um, e para a imagem quando ela é bloco solto. */
export function applyImageProperties(
  editor: Editor,
  placed: PlacedImage,
  properties: { readonly alt: string; readonly align: string | null },
): boolean {
  const chain = editor.chain().focus()

  chain.command(({ tr }) => {
    const node = tr.doc.nodeAt(placed.pos)
    if (node === null || node.type.name !== 'image') return false

    // Passo de atributo: `setNodeMarkup` substituiria a folha e a seleção se perderia.
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
