import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { OutlineBlock } from '@services/document/outline.js'
import { textBetweenWithoutNotes } from './extensions/note-ref.js'

/** Desce em tudo: o título numerado do Word é item de lista, e o de dentro de uma tabela também conta. */
export function outlineBlocksOf(doc: ProseMirrorNode): OutlineBlock[] {
  const blocks: OutlineBlock[] = []

  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    blocks.push({
      type: node.type.name,
      attrs: node.attrs,
      text: textBetweenWithoutNotes(node, 0, node.content.size),
      pos,
    })
    return false
  })

  return blocks
}
