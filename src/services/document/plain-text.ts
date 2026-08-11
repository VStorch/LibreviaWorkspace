import type { DocumentNode } from './model.js'

/**
 * Conversão entre o modelo do documento e texto puro.
 *
 * Existe para o `.txt` continuar sendo um formato de primeira classe: abrir um
 * arquivo de texto produz parágrafos comuns, e salvar como texto produz um
 * arquivo que qualquer programa lê. A perda de formatação nesse caminho é
 * inevitável — por isso `hasRichFormatting` avisa antes que ela aconteça.
 */

/** Nós que representam um bloco e, portanto, uma linha no texto puro. */
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
    // Células de tabela viram linhas próprias: sem colunas, é o melhor que o
    // texto puro consegue representar sem inventar alinhamento.
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

/**
 * O documento tem algo que o `.txt` não guarda?
 *
 * Serve para avisar antes de salvar, em vez de descobrir a perda depois. É
 * deliberadamente pessimista: na dúvida, avisa.
 */
export function hasRichFormatting(doc: DocumentNode): boolean {
  return anyNode(doc, (node) => {
    if (node.marks !== undefined && node.marks.length > 0) return true

    if (node.type !== 'doc' && node.type !== 'paragraph' && node.type !== 'text') return true

    // Parágrafo com alinhamento, recuo ou espaçamento também é formatação.
    const attrs = node.attrs
    if (attrs === undefined) return false
    return Object.values(attrs).some((value) => value !== null && value !== undefined && value !== 0)
  })
}

function anyNode(node: DocumentNode, predicate: (node: DocumentNode) => boolean): boolean {
  if (predicate(node)) return true
  return (node.content ?? []).some((child) => anyNode(child, predicate))
}
