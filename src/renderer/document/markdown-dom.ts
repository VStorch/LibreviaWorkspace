import { markdownToNodes } from '@services/document/markdown-import.js'
import type { DocumentNode } from '@services/document/model.js'
import { blockTagOf, linkOf, listStartOf, MARK_TAGS, type NodeMark } from './markdown-tags.js'

function withMark(content: Node, mark: NodeMark): Node {
  const tag = mark.type === 'link' ? 'span' : MARK_TAGS[mark.type]
  if (tag === undefined) return content
  const element = document.createElement(tag)
  if (mark.type === 'link') {
    element.className = 'ai-preview__link'
    element.title = linkOf(mark)
  }
  element.append(content)
  return element
}

/** Text goes in as text nodes: the answer never reaches the page as HTML. */
function domOf(node: DocumentNode): Node {
  if (node.type === 'text') {
    return (node.marks ?? []).reduce<Node>(withMark, document.createTextNode(node.text ?? ''))
  }
  if (node.type === 'hardBreak') return document.createElement('br')
  const element = document.createElement(blockTagOf(node))
  const start = listStartOf(node)
  if (start !== undefined) element.setAttribute('start', String(start))
  element.append(...(node.content ?? []).map(domOf))
  return element
}

/** As `MarkdownBlocks`, for a place outside React such as an editor widget. */
export function markdownDom(markdown: string): Node[] {
  return markdownToNodes(markdown).map(domOf)
}
