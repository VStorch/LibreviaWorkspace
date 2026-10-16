import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import {
  AllSelection,
  TextSelection,
  type EditorState,
  type Selection,
  type Transaction,
} from '@tiptap/pm/state'
import { exportMarkdown } from '@services/document/export-markdown.js'
import { markdownToNodes } from '@services/document/markdown-import.js'
import type { DocumentNode } from '@services/document/model.js'
import { plainPasteContent } from '@services/document/paste.js'
import { MAX_AI_DOCUMENT_LENGTH, MAX_AI_SELECTION_LENGTH } from '@shared/limits.js'
import { AiFormat } from '@shared/types.js'
import { DELETION, INSERTION, textWithoutDeletions } from './extensions/track-changes.js'
import { SKIP_TRACKING, trackTransaction } from './extensions/track-input.js'

export interface AiRange {
  readonly from: number
  readonly to: number
}

export type AiTargetProblem = 'empty' | 'unsupported' | 'tooLong'

export type AiTarget =
  { readonly range: AiRange; readonly text: string } | { readonly problem: AiTargetProblem }

/** Plain text in, plain text out: these would come back as letters, or not at all. */
const UNSUPPORTED_NODES: ReadonlySet<string> = new Set(['image', 'math', 'noteRef', 'field'])
const CELL_NODES: ReadonlySet<string> = new Set(['tableCell', 'tableHeader'])
/** Revisions are the tracking's. */
const UNCARRIED_MARKS: ReadonlySet<string> = new Set([INSERTION, DELETION])
/** Marks the Markdown answer says itself; the others (font, color, underline) come from the passage. */
const MARKDOWN_MARKS: ReadonlySet<string> = new Set(['bold', 'italic', 'strike', 'link'])
/** Exported as HTML or code that the answer would not bring back: they go as plain text. */
const PLAIN_IN_MARKDOWN: ReadonlySet<string> = new Set(['code', 'superscript', 'subscript'])
const TABLE_NODES: ReadonlySet<string> = new Set(['table', 'tableRow', ...CELL_NODES])

export interface AiResponse {
  readonly text: string
  readonly format: AiFormat
}

export function selectedText(doc: ProseMirrorNode, range: AiRange): string {
  return textWithoutDeletions(doc.cut(range.from, range.to), '\n', '')
}

function isDeletedNode(node: DocumentNode): boolean {
  return (node.marks ?? []).some((mark) => mark.type === DELETION)
}

/** A passage inside a cell goes as its paragraphs: a table would come back as a table's text. */
function markdownReady(nodes: readonly DocumentNode[]): DocumentNode[] {
  return nodes.flatMap((node): DocumentNode[] => {
    if (isDeletedNode(node)) return []
    const content = node.content === undefined ? undefined : markdownReady(node.content)
    if (TABLE_NODES.has(node.type)) return content ?? []
    const marks = node.marks?.filter((mark) => !PLAIN_IN_MARKDOWN.has(mark.type))
    return [
      { ...node, ...(content === undefined ? {} : { content }), ...(marks === undefined ? {} : { marks }) },
    ]
  })
}

export function selectedMarkdown(doc: ProseMirrorNode, range: AiRange): string {
  const content = markdownReady((doc.cut(range.from, range.to).toJSON() as DocumentNode).content ?? [])
  return exportMarkdown({ doc: { type: 'doc', content } }, { assetFolder: '' }).markdown.trimEnd()
}

/** The document around the selection, as far as the contract allows: the nearest text matters most. */
export function documentAround(
  doc: ProseMirrorNode,
  range: AiRange,
): { readonly before: string; readonly after: string } {
  const before = textWithoutDeletions(doc.cut(0, range.from), '\n', '')
  const after = textWithoutDeletions(doc.cut(range.to), '\n', '')
  return {
    before: before.slice(Math.max(0, before.length - MAX_AI_DOCUMENT_LENGTH)),
    after: after.slice(0, MAX_AI_DOCUMENT_LENGTH),
  }
}

export interface AiChatDocument {
  readonly before: string
  readonly after: string
  readonly selection: string
}

/** The document as it is now; a selection is marked only when it holds text that fits. */
export function chatDocument(state: EditorState): AiChatDocument {
  const selection = textualSelection(state)
  const range = { from: selection.from, to: selection.to }
  const text = selection.empty ? '' : selectedText(state.doc, range)
  if (text.trim() === '' || text.length > MAX_AI_SELECTION_LENGTH) {
    const whole = textWithoutDeletions(state.doc, '\n', '')
    return { before: whole.slice(0, MAX_AI_DOCUMENT_LENGTH), after: '', selection: '' }
  }
  return { ...documentAround(state.doc, range), selection: text }
}

function cellOf(state: EditorState, pos: number): number {
  const $pos = state.doc.resolve(pos)
  for (let depth = $pos.depth; depth > 0; depth--) {
    if (CELL_NODES.has($pos.node(depth).type.name)) return $pos.before(depth)
  }
  return -1
}

function hasUnsupportedNode(doc: ProseMirrorNode, range: AiRange): boolean {
  let found = false
  doc.nodesBetween(range.from, range.to, (node) => {
    if (UNSUPPORTED_NODES.has(node.type.name)) found = true
    return !found
  })
  return found
}

/** Ctrl+A selects the document node; the passage is the text from the first to the last letter. */
function textualSelection(state: EditorState): Selection {
  const { selection, doc } = state
  return selection instanceof AllSelection
    ? TextSelection.between(doc.resolve(0), doc.resolve(doc.content.size))
    : selection
}

export function aiTarget(state: EditorState): AiTarget {
  const selection = textualSelection(state)
  if (selection.empty) return { problem: 'empty' }
  const range = { from: selection.from, to: selection.to }
  const crossesCells = cellOf(state, range.from) !== cellOf(state, range.to)
  if (!(selection instanceof TextSelection) || crossesCells || hasUnsupportedNode(state.doc, range)) {
    return { problem: 'unsupported' }
  }
  const text = selectedText(state.doc, range)
  if (text.trim() === '') return { problem: 'empty' }
  if (text.length > MAX_AI_SELECTION_LENGTH) return { problem: 'tooLong' }
  return { range, text }
}

/** The passage edited since the request: applying would overwrite what the person typed. */
export function rangeChanged(doc: ProseMirrorNode, range: AiRange, snapshot: string): boolean {
  return range.to <= range.from || selectedText(doc, range) !== snapshot
}

function plainMarks(state: EditorState, pos: number, format: AiFormat): readonly Mark[] {
  return state.doc
    .resolve(pos)
    .marks()
    .filter(
      (mark) =>
        !UNCARRIED_MARKS.has(mark.type.name) &&
        (format === AiFormat.Plain || !MARKDOWN_MARKS.has(mark.type.name)),
    )
}

function paragraphsOf(state: EditorState, text: string, marks: readonly Mark[]): ProseMirrorNode[] {
  const { schema } = state
  return plainPasteContent(text).map((block) => {
    const line = block.type === 'text' ? (block.text ?? '') : (block.content?.[0]?.text ?? '')
    return schema.nodes['paragraph']!.create(null, line === '' ? null : schema.text(line, marks))
  })
}

/** The answer's own marks win over the passage's: `addToSet` replaces a mark of the same type. */
function withBaseMarks(node: ProseMirrorNode, base: readonly Mark[]): ProseMirrorNode {
  if (node.isText) return node.mark(node.marks.reduce((set, mark) => mark.addToSet(set), base))
  if (node.isTextblock && !node.type.allowsMarks(base)) return node
  const children: ProseMirrorNode[] = []
  node.forEach((child) => children.push(withBaseMarks(child, base)))
  return node.copy(Fragment.fromArray(children))
}

function blocksOf(state: EditorState, markdown: string, marks: readonly Mark[]): ProseMirrorNode[] {
  return markdownToNodes(markdown).map((block) => withBaseMarks(state.schema.nodeFromJSON(block), marks))
}

function responseBlocks(state: EditorState, response: AiResponse, marks: readonly Mark[]): ProseMirrorNode[] {
  return response.format === AiFormat.Markdown
    ? blocksOf(state, response.text, marks)
    : paragraphsOf(state, response.text, marks)
}

/**
 * Open paragraphs at both ends keep the change a plain replacement of text, which the revision
 * tracking follows; a list at the edge would wrap the text around it in a step it cannot mark.
 */
function textEdged(state: EditorState, blocks: readonly ProseMirrorNode[]): ProseMirrorNode[] {
  const empty = (): ProseMirrorNode[] =>
    blocks.length > 0 ? [state.schema.nodes['paragraph']!.create()] : []
  return [
    ...(blocks[0]?.isTextblock === true ? [] : empty()),
    ...blocks,
    ...(blocks[blocks.length - 1]?.isTextblock === true ? [] : empty()),
  ]
}

interface BlockParts {
  /** Lists and quotes that go whole before the passage's paragraph. */
  readonly leading: readonly ProseMirrorNode[]
  readonly middle: readonly ProseMirrorNode[]
  readonly trailing: readonly ProseMirrorNode[]
}

/**
 * A list at an edge of the passage with no text beyond it goes beside the paragraph, as whole
 * inserted blocks: an empty paragraph to hold it would remain after accepting or rejecting.
 */
function partition(blocks: readonly ProseMirrorNode[], headEmpty: boolean, tailEmpty: boolean): BlockParts {
  let start = 0
  let end = blocks.length
  if (headEmpty) while (start < end && !blocks[start]!.isTextblock) start++
  if (tailEmpty) while (end > start && !blocks[end - 1]!.isTextblock) end--
  return { leading: blocks.slice(0, start), middle: blocks.slice(start, end), trailing: blocks.slice(end) }
}

function replaceInParagraphs(
  tr: Transaction,
  state: EditorState,
  range: AiRange,
  blocks: readonly ProseMirrorNode[],
): Transaction {
  const $from = tr.doc.resolve(range.from)
  const $to = tr.doc.resolve(range.to)
  const parts = partition(blocks, $from.parentOffset === 0, $to.parentOffset === $to.parent.content.size)
  if (parts.trailing.length > 0) tr.insert($to.after(), Fragment.fromArray(parts.trailing))
  const middle = parts.middle.length > 0 ? textEdged(state, parts.middle) : []
  tr.replace(
    range.from,
    range.to,
    middle.length > 0 ? new Slice(Fragment.fromArray(middle), 1, 1) : Slice.empty,
  )
  if (parts.leading.length > 0) tr.insert($from.before(), Fragment.fromArray(parts.leading))
  return tr
}

function tracked(tr: Transaction, state: EditorState, author: string, now: Date): Transaction {
  return trackTransaction(tr, state, author, now).setMeta(SKIP_TRACKING, true)
}

/** Lines become paragraphs, the first and last joined to the text around the passage, as a paste. */
export function buildTrackedReplacement(
  state: EditorState,
  range: AiRange,
  response: AiResponse,
  author: string,
  now: Date,
): Transaction | null {
  const blocks = responseBlocks(state, response, plainMarks(state, range.from, response.format))
  if (blocks.length === 0) return null
  return tracked(replaceInParagraphs(state.tr, state, range, blocks), state, author, now)
}

export interface AiTextEdit {
  readonly range: AiRange
  readonly text: string
}

/**
 * Several plain-text replacements as one tracked change, one undo. The ranges must not overlap;
 * the last goes first so the earlier positions still hold.
 */
export function buildTrackedEdits(
  state: EditorState,
  edits: readonly AiTextEdit[],
  author: string,
  now: Date,
): Transaction | null {
  if (edits.length === 0) return null
  const tr = state.tr
  for (const edit of [...edits].sort((a, b) => b.range.from - a.range.from)) {
    const marks = plainMarks(state, edit.range.from, AiFormat.Plain)
    const blocks = paragraphsOf(state, edit.text, marks)
    replaceInParagraphs(tr, state, edit.range, edit.text === '' ? [] : blocks)
  }
  return tracked(tr, state, author, now)
}

export function buildInsertBelow(
  state: EditorState,
  range: AiRange,
  response: AiResponse,
  author: string,
  now: Date,
): Transaction | null {
  const blocks = responseBlocks(state, response, [])
  if (blocks.length === 0) return null
  const $to = state.doc.resolve(range.to)
  const after = $to.parent.isTextblock ? $to.after() : range.to
  return tracked(state.tr.insert(after, Fragment.fromArray(blocks)), state, author, now)
}
