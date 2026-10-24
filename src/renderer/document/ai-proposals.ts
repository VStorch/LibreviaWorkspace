import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { occurrences, type AiAnchoredProposal, type AiInsertProposal } from '@services/ai/proposals.js'
import { MAX_AI_DOCUMENT_LENGTH, MAX_AI_PARAGRAPHS } from '@shared/limits.js'
import type { AiRange } from './ai-apply.js'
import { isDeleted } from './extensions/track-changes.js'

/** Where a stretch of the paragraph's text sits in the document. */
interface TextSpan {
  readonly offset: number
  readonly pos: number
  readonly length: number
}

interface ParagraphText {
  readonly text: string
  readonly spans: readonly TextSpan[]
  readonly pos: number
  readonly node: ProseMirrorNode
}

const NOTE_REF = 'noteRef'

/** Deleted text is left out; comment ends, breaks and other leaves add nothing. */
function paragraphText(block: ProseMirrorNode, blockPos: number): ParagraphText {
  const spans: TextSpan[] = []
  let text = ''
  block.forEach((child, childOffset) => {
    if (!child.isText || isDeleted(child)) return
    const piece = child.text ?? ''
    spans.push({ offset: text.length, pos: blockPos + 1 + childOffset, length: piece.length })
    text += piece
  })
  return { text, spans, pos: blockPos, node: block }
}

/**
 * Every text block in document order, cells and list items included, notes left out: the numbering
 * the model saw, which applying a proposal walks again.
 */
function paragraphsOf(doc: ProseMirrorNode): ParagraphText[] {
  const paragraphs: ParagraphText[] = []
  doc.descendants((node, pos) => {
    if (node.type.name === NOTE_REF) return false
    if (!node.isTextblock) return true
    paragraphs.push(paragraphText(node, pos))
    return false
  })
  return paragraphs
}

/** Whole paragraphs from the first, as many as the contract allows: a cut one would misplace a find. */
export function chatParagraphs(doc: ProseMirrorNode): string[] {
  const texts: string[] = []
  let total = 0
  for (const { text } of paragraphsOf(doc)) {
    if (texts.length >= MAX_AI_PARAGRAPHS || total + text.length > MAX_AI_DOCUMENT_LENGTH) break
    texts.push(text)
    total += text.length
  }
  return texts
}

function startPos(spans: readonly TextSpan[], offset: number): number | null {
  const span = spans.find(
    (candidate) => offset >= candidate.offset && offset < candidate.offset + candidate.length,
  )
  return span === undefined ? null : span.pos + offset - span.offset
}

function endPos(spans: readonly TextSpan[], offset: number): number | null {
  const span = spans.find(
    (candidate) => offset > candidate.offset && offset <= candidate.offset + candidate.length,
  )
  return span === undefined ? null : span.pos + offset - span.offset
}

/**
 * Null when the paragraph of that number no longer reads as it did: an edit or a paragraph added
 * above would otherwise land the proposal on another one.
 */
export function locateProposal(doc: ProseMirrorNode, proposal: AiAnchoredProposal): AiRange | null {
  const paragraph = paragraphsOf(doc)[proposal.paragraph - 1]
  if (paragraph?.text !== proposal.paragraphText) return null
  if (occurrences(paragraph.text, proposal.find) !== 1) return null
  const offset = paragraph.text.indexOf(proposal.find)
  const from = startPos(paragraph.spans, offset)
  const to = endPos(paragraph.spans, offset + proposal.find.length)
  return from === null || to === null ? null : { from, to }
}

/**
 * After the top-level block that holds the paragraph, so a list or table stays whole; an empty
 * paragraph of its own is the place to write, so the insertion goes before it. Null as in
 * `locateProposal`.
 */
export function locateInsertion(doc: ProseMirrorNode, proposal: AiInsertProposal): number | null {
  if (proposal.paragraph === 0) return 0
  const paragraph = paragraphsOf(doc)[proposal.paragraph - 1]
  if (paragraph?.text !== proposal.paragraphText) return null
  const $block = doc.resolve(paragraph.pos)
  if ($block.depth > 0) return $block.after(1)
  return paragraph.node.content.size === 0 ? paragraph.pos : paragraph.pos + paragraph.node.nodeSize
}

export interface LocatedProposal<P extends AiAnchoredProposal> {
  readonly proposal: P
  readonly range: AiRange
}

export interface LocatedProposals<P extends AiAnchoredProposal> {
  readonly located: readonly LocatedProposal<P>[]
  /** Not found as they were, or overlapping one found before them. */
  readonly changed: readonly string[]
}

const overlaps = (a: AiRange, b: AiRange): boolean => a.from < b.to && b.from < a.to

export function locateProposals<P extends AiAnchoredProposal>(
  doc: ProseMirrorNode,
  proposals: readonly P[],
): LocatedProposals<P> {
  const located: LocatedProposal<P>[] = []
  const changed: string[] = []
  for (const proposal of proposals) {
    const range = locateProposal(doc, proposal)
    if (range === null || located.some((other) => overlaps(other.range, range))) changed.push(proposal.id)
    else located.push({ proposal, range })
  }
  return { located, changed }
}
