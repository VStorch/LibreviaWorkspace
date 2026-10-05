import Image from '@tiptap/extension-image'
import { mergeAttributes } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { ReactNodeViewRenderer } from '@tiptap/react'
import { ImageNodeView } from '../ImageNodeView.js'

/**
 * `wp:extent` need not have the file's ratio: Word draws it stretched. `width` and `height` alone
 * would become `aspect-ratio: auto`, and the file's ratio would win; declared without `auto`, the
 * box is reserved before the image decodes, which is when pagination measures. The `NodeView` makes
 * room for the handles; `renderHTML` applies when printing.
 */
export const DocumentImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),

      /**
       * The paragraph anchoring the image takes its own line besides it, as LibreOffice measures.
       */
      anchored: {
        default: null,
        parseHTML: (element: HTMLElement) => (element.hasAttribute('data-anchored') ? true : null),
        renderHTML: (attributes: Record<string, unknown>) =>
          attributes['anchored'] === true ? { 'data-anchored': '' } : {},
      },

      /** Only for block images in old drafts: in OOXML the paragraph's `w:jc` aligns. */
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
      // On the outer element, the paragraph's child: that is what the stylesheet rules look at.
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
  /** It aligns an image read from the file. */
  readonly paragraphPos: number | null
}

/**
 * Selected, or in the cursor's block: arriving with the arrow keys leaves the cursor in the
 * paragraph.
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

function isAligned(block: ProseMirrorNode): boolean {
  return block.type.name === 'paragraph' || block.type.name === 'heading'
}

/**
 * Alignment goes to the **paragraph** when there is one, and to the image when it is a loose block.
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

    // An attribute step: `setNodeMarkup` would replace the leaf and the selection would be lost.
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
