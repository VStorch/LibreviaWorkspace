import { Lexer, type MarkedToken, type Token, type Tokens } from 'marked'
import { normalizeLinkUrl } from './link.js'
import type { DocumentNode } from './model.js'

/**
 * Markdown from an AI model into document nodes. Only the lexer: marked's HTML renderer never runs,
 * and raw HTML in the answer stays as the literal text the model wrote. Code comes in as plain
 * text: the editor's code mark and code block refuse the revision mark the answer must carry.
 */

type NodeMark = NonNullable<DocumentNode['marks']>[number]

const BOLD: NodeMark = { type: 'bold' }
const ITALIC: NodeMark = { type: 'italic' }
const STRIKE: NodeMark = { type: 'strike' }
const SOFT_BREAK = /\n/g

export function markdownToNodes(markdown: string): DocumentNode[] {
  return blocksOf(new Lexer({ gfm: true }).lex(markdown))
}

function blocksOf(tokens: readonly Token[]): DocumentNode[] {
  return tokens.flatMap((token) => blockOf(token as MarkedToken))
}

function blockOf(token: MarkedToken): DocumentNode[] {
  switch (token.type) {
    case 'paragraph':
    case 'text':
      return [textblock('paragraph', inlineOfBlock(token))]
    case 'heading':
      return [{ ...textblock('heading', inlineOf(token.tokens, [])), attrs: { level: token.depth } }]
    case 'list':
      return [listOf(token)]
    case 'blockquote':
      return [{ type: 'blockquote', content: atLeastOneBlock(blocksOf(token.tokens)) }]
    case 'code':
      return literalParagraphs(token.text)
    case 'space':
    case 'def':
    case 'hr':
      return []
    default:
      return literalParagraphs(token.raw)
  }
}

function inlineOfBlock(token: Tokens.Paragraph | Tokens.Text): DocumentNode[] {
  return token.tokens === undefined ? textOf(softBreaks(token.text), []) : inlineOf(token.tokens, [])
}

function textblock(type: string, content: DocumentNode[]): DocumentNode {
  return content.length === 0 ? { type } : { type, content }
}

function listOf(token: Tokens.List): DocumentNode {
  const items = token.items.map((item) => ({
    type: 'listItem',
    content: atLeastOneBlock(blocksOf(item.tokens)),
  }))
  if (!token.ordered) return { type: 'bulletList', content: items }
  const start = typeof token.start === 'number' ? token.start : 1
  return { type: 'orderedList', attrs: { start }, content: items }
}

/** A list item and a quote must begin with a paragraph. */
function atLeastOneBlock(blocks: DocumentNode[]): DocumentNode[] {
  return blocks[0]?.type === 'paragraph' ? blocks : [{ type: 'paragraph' }, ...blocks]
}

/** A table or anything else without a counterpart here keeps the text the model wrote. */
function literalParagraphs(raw: string): DocumentNode[] {
  return raw
    .trim()
    .split('\n')
    .filter((line) => line.trim() !== '')
    .map((line) => textblock('paragraph', textOf(line, [])))
}

function inlineOf(tokens: readonly Token[], marks: readonly NodeMark[]): DocumentNode[] {
  return tokens.flatMap((token) => inlineToken(token as MarkedToken, marks))
}

function inlineToken(token: MarkedToken, marks: readonly NodeMark[]): DocumentNode[] {
  switch (token.type) {
    case 'strong':
      return inlineOf(token.tokens, withMark(marks, BOLD))
    case 'em':
      return inlineOf(token.tokens, withMark(marks, ITALIC))
    case 'del':
      return inlineOf(token.tokens, withMark(marks, STRIKE))
    case 'codespan':
      return textOf(token.text, marks)
    case 'link':
      return inlineOf(token.tokens, withLink(marks, token.href))
    case 'br':
      return [{ type: 'hardBreak' }]
    case 'text':
      return token.tokens === undefined
        ? textOf(softBreaks(token.text), marks)
        : inlineOf(token.tokens, marks)
    case 'escape':
    case 'image':
      return textOf(token.text, marks)
    default:
      return textOf(token.raw, marks)
  }
}

/** An address the editor would refuse (`javascript:`) leaves the words without a link. */
function withLink(marks: readonly NodeMark[], href: string): readonly NodeMark[] {
  const safe = normalizeLinkUrl(href)
  return safe === null ? marks : withMark(marks, { type: 'link', attrs: { href: safe } })
}

/** `**a **b** c**` would otherwise carry bold twice. */
function withMark(marks: readonly NodeMark[], mark: NodeMark): readonly NodeMark[] {
  return [...marks.filter((other) => other.type !== mark.type), mark]
}

function softBreaks(text: string): string {
  return text.replace(SOFT_BREAK, ' ')
}

function textOf(text: string, marks: readonly NodeMark[]): DocumentNode[] {
  if (text === '') return []
  return [marks.length === 0 ? { type: 'text', text } : { type: 'text', text, marks: [...marks] }]
}
