import type { DocumentNode } from './model.js'

const BLOCK_TYPES = new Set([
  'paragraph',
  'heading',
  'blockquote',
  'listItem',
  'codeBlock',
  'tableCell',
  'tableHeader',
])

export function plainTextToDocument(text: string): DocumentNode {
  const lines = text.split(/\r\n|\r|\n/)
  const paragraphs = lines.map<DocumentNode>((line) =>
    line.length === 0
      ? { type: 'paragraph' }
      : { type: 'paragraph', content: [{ type: 'text', text: line }] },
  )

  // Um documento do ProseMirror não pode ser vazio.
  return { type: 'doc', content: paragraphs.length > 0 ? paragraphs : [{ type: 'paragraph' }] }
}

export function documentToPlainText(doc: DocumentNode): string {
  const lines: string[] = []
  collectLines(doc, lines)
  return lines.join('\n')
}

function collectLines(node: DocumentNode, lines: string[]): void {
  if (BLOCK_TYPES.has(node.type)) {
    const text = collectInlineText(node)
    // Sem colunas, cada célula vira uma linha.
    lines.push(text)
    if (node.type === 'tableCell' || node.type === 'tableHeader') return
  }

  if (node.type === 'horizontalRule' || node.type === 'pageBreak') {
    lines.push('')
    return
  }

  for (const child of node.content ?? []) {
    if (BLOCK_TYPES.has(node.type) && !hasBlockDescendant(child)) continue
    collectLines(child, lines)
  }
}

function hasBlockDescendant(node: DocumentNode): boolean {
  if (BLOCK_TYPES.has(node.type)) return true
  return (node.content ?? []).some(hasBlockDescendant)
}

function collectInlineText(node: DocumentNode): string {
  if (node.type === 'text') return node.text ?? ''
  if (node.type === 'hardBreak') return '\n'
  if (BLOCK_TYPES.has(node.type) && node !== undefined) {
    return (node.content ?? [])
      .filter((child) => !hasBlockDescendant(child))
      .map(collectInlineText)
      .join('')
  }
  return (node.content ?? []).map(collectInlineText).join('')
}

/** Pessimista: na dúvida, avisa antes de salvar. */
export function hasRichFormatting(doc: DocumentNode): boolean {
  return anyNode(doc, (node) => {
    if (node.marks !== undefined && node.marks.length > 0) return true

    if (node.type !== 'doc' && node.type !== 'paragraph' && node.type !== 'text') return true

    const attrs = node.attrs
    if (attrs === undefined) return false
    return Object.values(attrs).some((value) => value !== null && value !== undefined && value !== 0)
  })
}

function anyNode(node: DocumentNode, predicate: (node: DocumentNode) => boolean): boolean {
  if (predicate(node)) return true
  return (node.content ?? []).some((child) => anyNode(child, predicate))
}
