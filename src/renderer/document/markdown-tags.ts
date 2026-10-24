import type { DocumentNode } from '@services/document/model.js'

export type NodeMark = NonNullable<DocumentNode['marks']>[number]

const BLOCK_TAGS: Readonly<Record<string, string>> = {
  paragraph: 'p',
  blockquote: 'blockquote',
  bulletList: 'ul',
  orderedList: 'ol',
  listItem: 'li',
  codeBlock: 'pre',
}

export const MARK_TAGS: Readonly<Record<string, string>> = {
  bold: 'strong',
  italic: 'em',
  strike: 's',
  code: 'code',
}

const HEADING_TAGS: readonly string[] = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']

export function blockTagOf(node: DocumentNode): string {
  if (node.type !== 'heading') return BLOCK_TAGS[node.type] ?? 'p'
  const level = Number(node.attrs?.['level']) || 1
  return HEADING_TAGS[Math.min(HEADING_TAGS.length, Math.max(1, level)) - 1]!
}

export const listStartOf = (node: DocumentNode): number | undefined =>
  node.type === 'orderedList' ? Number(node.attrs?.['start'] ?? 1) : undefined

export const linkOf = (mark: NodeMark): string => String(mark.attrs?.['href'] ?? '')
