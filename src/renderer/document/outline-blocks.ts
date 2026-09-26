import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { OutlineBlock } from '@services/document/outline.js'

/**
 * Os blocos de texto do documento, na ordem, com a posição de cada um.
 *
 * Desce em tudo — item de lista, célula, citação — porque o título numerado do
 * Word é item de lista, e o de dentro de uma tabela também é título para ele. O
 * bloco de texto é a folha da descida: dentro dele só há linha.
 */
export function outlineBlocksOf(doc: ProseMirrorNode): OutlineBlock[] {
  const blocks: OutlineBlock[] = []

  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    blocks.push({ type: node.type.name, attrs: node.attrs, text: node.textContent, pos })
    return false
  })

  return blocks
}
