import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import {
  AiProposalKind,
  type AiCommentProposal,
  type AiEditProposal,
  type AiInsertProposal,
  type AiProposal,
} from '@services/ai/proposals.js'
import { DiffOp, wordDiff } from '@services/ai/word-diff.js'
import { markdownToNodes } from '@services/document/markdown-import.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { t } from '../../i18n.js'
import { edgeSpacesOf, type AiRange } from '../ai-apply.js'
import { placeProposals } from '../ai-proposals.js'
import { markdownDom } from '../markdown-dom.js'

export interface AiDiffEntry {
  /** Unique among all the open proposals of the conversation. */
  readonly key: string
  readonly proposal: AiProposal
}

export interface AiDiffInput {
  readonly entries: readonly AiDiffEntry[]
  readonly readOnly: boolean
  readonly decide: (key: string, accept: boolean) => void
}

interface AiDiffState extends AiDiffInput {
  readonly decorations: DecorationSet
}

export const aiDiffKey = new PluginKey<AiDiffState>('aiDiff')

const EMPTY: AiDiffState = {
  entries: [],
  readOnly: false,
  decide: () => undefined,
  decorations: DecorationSet.empty,
}

const ACCEPT_LABELS: Readonly<Record<AiProposalKind, MessageKey>> = {
  [AiProposalKind.Edit]: 'ai.diff.acceptEdit',
  [AiProposalKind.Insert]: 'ai.diff.acceptInsert',
  [AiProposalKind.Comment]: 'ai.diff.acceptComment',
}

const REJECT_LABELS: Readonly<Record<AiProposalKind, MessageKey>> = {
  [AiProposalKind.Edit]: 'ai.diff.rejectEdit',
  [AiProposalKind.Insert]: 'ai.diff.rejectInsert',
  [AiProposalKind.Comment]: 'ai.diff.rejectComment',
}

/** Outside the history: showing a proposal is not an edit the person would undo. */
export function showAiDiff(view: EditorView, input: AiDiffInput): void {
  view.dispatch(view.state.tr.setMeta(aiDiffKey, input).setMeta('addToHistory', false))
}

function element(tag: string, className: string, children: readonly (Node | string)[] = []): HTMLElement {
  const node = document.createElement(tag)
  node.className = className
  node.append(...children)
  return node
}

function decisionButton(
  view: EditorView,
  entry: AiDiffEntry,
  accept: boolean,
  readOnly: boolean,
): HTMLElement {
  const button = element('button', accept ? 'ai-diff__accept' : 'ai-diff__reject', [
    t(accept ? 'ai.diff.accept' : 'ai.diff.reject'),
  ]) as HTMLButtonElement
  button.type = 'button'
  button.disabled = accept && readOnly
  button.setAttribute('aria-label', t((accept ? ACCEPT_LABELS : REJECT_LABELS)[entry.proposal.kind]))
  // Keeps the caret where it was: the click decides, it does not move the selection; the text
  // takes the focus after, so undo is at hand.
  button.addEventListener('mousedown', (event) => event.preventDefault())
  button.addEventListener('click', () => {
    aiDiffKey.getState(view.state)?.decide(entry.key, accept)
    view.focus()
  })
  return button
}

function actions(view: EditorView, entry: AiDiffEntry, readOnly: boolean): HTMLElement {
  const bar = element('span', 'ai-diff__actions', [
    decisionButton(view, entry, true, readOnly),
    decisionButton(view, entry, false, readOnly),
  ])
  bar.contentEditable = 'false'
  return bar
}

const widgetSpec = (key: string, side: number) => ({
  key,
  side,
  ignoreSelection: true,
  stopEvent: () => true,
})

function actionsWidget(pos: number, entry: AiDiffEntry, readOnly: boolean): Decoration {
  return Decoration.widget(
    pos,
    (view) => actions(view, entry, readOnly),
    widgetSpec(`${entry.key}:actions:${readOnly}`, 1),
  )
}

function blockWidget(pos: number, entry: AiDiffEntry, markdown: string, readOnly: boolean): Decoration {
  return Decoration.widget(
    pos,
    (view) => {
      const block = element('div', 'ai-diff__block', markdownDom(markdown))
      block.contentEditable = 'false'
      block.append(actions(view, entry, readOnly))
      return block
    },
    widgetSpec(`${entry.key}:block:${markdown}:${readOnly}`, 1),
  )
}

/** The new text when the answer is one plain paragraph, which a word diff can show inline. */
function inlineText(markdown: string): string | null {
  if (markdown === '') return ''
  const { lead, core, trail } = edgeSpacesOf(markdown)
  const blocks = core === '' ? [] : markdownToNodes(core)
  if (blocks.length > 1 || (blocks[0] !== undefined && blocks[0].type !== 'paragraph')) return null
  const pieces = (blocks[0]?.content ?? []).map((node) =>
    node.type === 'hardBreak' ? '\n' : (node.text ?? ''),
  )
  return lead + pieces.join('') + trail
}

function wordDecorations(entry: AiDiffEntry, find: string, range: AiRange, after: string): Decoration[] {
  const decorations: Decoration[] = []
  let at = range.from
  wordDiff(find, after).forEach((piece, index) => {
    if (piece.op === DiffOp.Added) {
      const added = (): HTMLElement => element('span', 'ai-diff__new', [piece.text])
      decorations.push(Decoration.widget(at, added, widgetSpec(`${entry.key}:${index}:${piece.text}`, -1)))
      return
    }
    if (piece.op === DiffOp.Removed)
      decorations.push(Decoration.inline(at, at + piece.text.length, { class: 'ai-diff__old' }))
    at += piece.text.length
  })
  return decorations
}

/** After the top-level block that holds the range: a widget inside a paragraph would break its lines. */
function blockEnd(doc: ProseMirrorNode, range: AiRange): number {
  const $to = doc.resolve(range.to)
  return $to.depth === 0 ? range.to : $to.after(1)
}

function editDecorations(
  doc: ProseMirrorNode,
  entry: AiDiffEntry,
  proposal: AiEditProposal,
  range: AiRange,
  readOnly: boolean,
): Decoration[] {
  const after = inlineText(proposal.replace)
  // Deleted text or an equation inside the range would shift the diff off the letters it marks.
  const plain = doc.textBetween(range.from, range.to, '\n', '￼') === proposal.find
  if (after !== null && plain) {
    return [...wordDecorations(entry, proposal.find, range, after), actionsWidget(range.to, entry, readOnly)]
  }
  const old = Decoration.inline(range.from, range.to, { class: 'ai-diff__old' })
  if (proposal.replace === '') return [old, actionsWidget(range.to, entry, readOnly)]
  return [old, blockWidget(blockEnd(doc, range), entry, proposal.replace, readOnly)]
}

function commentDecorations(
  entry: AiDiffEntry,
  proposal: AiCommentProposal,
  range: AiRange,
  readOnly: boolean,
): Decoration[] {
  const note = (view: EditorView): HTMLElement => {
    const span = element('span', 'ai-diff__note', [proposal.comment])
    span.contentEditable = 'false'
    span.append(actions(view, entry, readOnly))
    return span
  }
  return [
    Decoration.inline(range.from, range.to, { class: 'ai-diff__comment' }),
    Decoration.widget(range.to, note, widgetSpec(`${entry.key}:note:${proposal.comment}:${readOnly}`, 1)),
  ]
}

function insertDecorations(
  entry: AiDiffEntry,
  proposal: AiInsertProposal,
  pos: number,
  readOnly: boolean,
): Decoration[] {
  return [blockWidget(pos, entry, proposal.content, readOnly)]
}

function buildDecorations(doc: ProseMirrorNode, input: AiDiffInput): DecorationSet {
  const places = placeProposals(
    doc,
    input.entries.map(({ proposal }) => proposal),
  )
  const decorations = input.entries.flatMap((entry, index): Decoration[] => {
    const place = places[index]
    if (place === null || place === undefined) return []
    const { proposal } = entry
    if (proposal.kind === AiProposalKind.Insert) {
      return typeof place === 'number' ? insertDecorations(entry, proposal, place, input.readOnly) : []
    }
    if (typeof place === 'number') return []
    return proposal.kind === AiProposalKind.Edit
      ? editDecorations(doc, entry, proposal, place, input.readOnly)
      : commentDecorations(entry, proposal, place, input.readOnly)
  })
  return decorations.length === 0 ? DecorationSet.empty : DecorationSet.create(doc, decorations)
}

export function aiDiffPlugin(): Plugin<AiDiffState> {
  return new Plugin<AiDiffState>({
    key: aiDiffKey,
    state: {
      init: () => EMPTY,
      apply(tr, value, _oldState, newState) {
        const input = tr.getMeta(aiDiffKey) as AiDiffInput | undefined
        if (input === undefined && !tr.docChanged) return value
        const next = input ?? value
        if (next.entries.length === 0) return { ...next, decorations: DecorationSet.empty }
        return { ...next, decorations: buildDecorations(newState.doc, next) }
      },
    },
    props: {
      decorations: (state) => aiDiffKey.getState(state)?.decorations ?? DecorationSet.empty,
    },
  })
}

/**
 * The chat's open proposals drawn on the sheet, as a diff: what goes struck through, what comes in
 * beside it, and the buttons to accept or reject each one. Nothing here changes the document.
 */
export const AiDiff = Extension.create({
  name: 'aiDiff',

  addProseMirrorPlugins() {
    return [aiDiffPlugin()]
  },
})
