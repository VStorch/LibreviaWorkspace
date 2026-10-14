import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import {
  AllSelection,
  TextSelection,
  type EditorState,
  type Selection,
  type Transaction,
} from '@tiptap/pm/state'
import { plainPasteContent } from '@services/document/paste.js'
import { MAX_AI_SELECTION_LENGTH } from '@shared/limits.js'
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
const REVISION_MARKS: ReadonlySet<string> = new Set([INSERTION, DELETION])

export function selectedText(doc: ProseMirrorNode, range: AiRange): string {
  return textWithoutDeletions(doc.cut(range.from, range.to), '\n', '')
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

function plainMarks(state: EditorState, pos: number): readonly Mark[] {
  return state.doc
    .resolve(pos)
    .marks()
    .filter((mark) => !REVISION_MARKS.has(mark.type.name))
}

function paragraphsOf(state: EditorState, text: string, marks: readonly Mark[]): ProseMirrorNode[] {
  const { schema } = state
  return plainPasteContent(text).map((block) => {
    const line = block.type === 'text' ? (block.text ?? '') : (block.content?.[0]?.text ?? '')
    return schema.nodes['paragraph']!.create(null, line === '' ? null : schema.text(line, marks))
  })
}

function tracked(tr: Transaction, state: EditorState, author: string, now: Date): Transaction {
  return trackTransaction(tr, state, author, now).setMeta(SKIP_TRACKING, true)
}

/** Lines become paragraphs, the first and last joined to the text around the passage, as a paste. */
export function buildTrackedReplacement(
  state: EditorState,
  range: AiRange,
  text: string,
  author: string,
  now: Date,
): Transaction | null {
  const paragraphs = paragraphsOf(state, text, plainMarks(state, range.from))
  if (paragraphs.length === 0) return null
  const slice =
    paragraphs.length === 1
      ? new Slice(paragraphs[0]!.content, 0, 0)
      : new Slice(Fragment.fromArray(paragraphs), 1, 1)
  return tracked(state.tr.replace(range.from, range.to, slice), state, author, now)
}

export function buildInsertBelow(
  state: EditorState,
  range: AiRange,
  text: string,
  author: string,
  now: Date,
): Transaction | null {
  const paragraphs = paragraphsOf(state, text, [])
  if (paragraphs.length === 0) return null
  const $to = state.doc.resolve(range.to)
  const after = $to.parent.isTextblock ? $to.after() : range.to
  return tracked(state.tr.insert(after, Fragment.fromArray(paragraphs)), state, author, now)
}
