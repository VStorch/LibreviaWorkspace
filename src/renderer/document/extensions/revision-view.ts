import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode, ResolvedPos } from '@tiptap/pm/model'
import {
  Plugin,
  PluginKey,
  Selection,
  TextSelection,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import { RevisionView } from '@shared/types.js'
import { DELETION, INSERTION, ZERO_WIDTH, blockRevisionOf, characterSize } from './track-changes.js'

/**
 * All markup: everything visible. Simple: the final text, with a bar in the margin of the changed
 * paragraph. No markup: the final text, clean. Original: the earlier text, with the editor locked.
 * CSS hides, through the `revisions-<mode>` class; the document does not change. The cursor does
 * not land in hidden text: the selection is pushed to the edge, and Backspace and Delete skip over
 * it.
 */

interface ViewState {
  readonly view: RevisionView
  readonly decorations: DecorationSet
}

export const revisionViewKey = new PluginKey<ViewState>('revisionView')

export function revisionViewOf(state: EditorState): RevisionView {
  return revisionViewKey.getState(state)?.view ?? RevisionView.All
}

/** Without changing the document. */
export function setRevisionViewMeta(tr: Transaction, view: RevisionView): Transaction {
  return tr.setMeta(revisionViewKey, view)
}

export function isHiddenInline(node: ProseMirrorNode, view: RevisionView): boolean {
  if (view === RevisionView.All) return false
  const name = view === RevisionView.Original ? INSERTION : DELETION
  return node.marks.some((mark) => mark.type.name === name)
}

/** A paragraph mark or a row. */
function isHiddenRevision(value: unknown, view: RevisionView): boolean {
  if (view === RevisionView.All) return false
  const kind = blockRevisionOf(value)?.kind
  return view === RevisionView.Original ? kind === 'ins' : kind === 'del'
}

/** And there is something to hide. */
function allContentHidden(block: ProseMirrorNode, view: RevisionView): boolean {
  let hidden = false
  let visible = false
  block.forEach((child) => {
    if (isHiddenInline(child, view)) hidden = true
    else if (!ZERO_WIDTH.has(child.type.name)) visible = true
  })
  return hidden && !visible
}

/**
 * A table row of a hidden revision, and a paragraph whose mark disappears along with all its text.
 */
export function isHiddenBlock(node: ProseMirrorNode, view: RevisionView): boolean {
  if (view === RevisionView.All) return false
  if (node.type.name === 'tableRow') return isHiddenRevision(node.attrs['rowRevision'], view)
  if (!node.isTextblock || !isHiddenRevision(node.attrs['markRevision'], view)) return false
  return node.content.size === 0 || allContentHidden(node, view)
}

/** In block-relative positions, crossing zero-width ends between hidden runs. */
export function hiddenRuns(block: ProseMirrorNode, view: RevisionView): { from: number; to: number }[] {
  const runs: { from: number; to: number }[] = []
  if (view === RevisionView.All || !block.isTextblock) return runs
  let run: { from: number; to: number } | null = null
  block.forEach((child, offset) => {
    if (isHiddenInline(child, view)) {
      if (run === null) {
        run = { from: offset, to: offset + child.nodeSize }
        runs.push(run)
      } else {
        run.to = offset + child.nodeSize
      }
    } else if (!ZERO_WIDTH.has(child.type.name)) {
      run = null
    }
  })
  return runs
}

/** In direction `dir`, or the other way when there is no exit. */
export function visiblePosition(doc: ProseMirrorNode, pos: number, view: RevisionView, dir: 1 | -1): number {
  if (view === RevisionView.All) return pos
  // The seesaw between hidden blocks is recognized by a repeated position, not by a loop count.
  const seen = new Set<number>()
  while (!seen.has(pos)) {
    seen.add(pos)
    const $pos = doc.resolve(pos)
    const outside = outOfHiddenBlock(doc, $pos, view, dir)
    if (outside === null) return pos
    if (outside !== pos) {
      pos = outside
      continue
    }
    const start = $pos.start()
    for (const run of hiddenRuns($pos.parent, view)) {
      if (run.from < $pos.parentOffset && $pos.parentOffset < run.to) {
        return start + (dir > 0 ? run.to : run.from)
      }
    }
    return pos
  }
  return pos
}

/** The same one if there is none, or `null` when there is no exit. */
function outOfHiddenBlock(
  doc: ProseMirrorNode,
  $pos: ResolvedPos,
  view: RevisionView,
  dir: 1 | -1,
): number | null {
  for (let depth = 1; depth <= $pos.depth; depth++) {
    if (!isHiddenBlock($pos.node(depth), view)) continue
    const edge = dir > 0 ? $pos.after(depth) : $pos.before(depth)
    const found =
      Selection.findFrom(doc.resolve(edge), dir, true) ??
      Selection.findFrom(doc.resolve(dir > 0 ? $pos.before(depth) : $pos.after(depth)), -dir as 1 | -1, true)
    return found === null ? null : found.head
  }
  return $pos.pos
}

/**
 * Hidden deleted text in the middle is not erased by a selection that passed over it without
 * showing it.
 */
export function deleteVisible(tr: Transaction, from: number, to: number, view: RevisionView): Transaction {
  const hidden: { from: number; to: number }[] = []
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isInline && isHiddenInline(node, view)) {
      hidden.push({ from: Math.max(pos, from), to: Math.min(pos + node.nodeSize, to) })
    }
    return true
  })
  const pieces: { from: number; to: number }[] = []
  let cursor = from
  for (const range of hidden) {
    if (range.from > cursor) pieces.push({ from: cursor, to: range.from })
    cursor = Math.max(cursor, range.to)
  }
  if (cursor < to) pieces.push({ from: cursor, to })
  for (const piece of pieces.reverse()) tr.delete(piece.from, piece.to)
  return tr
}

/** `null` if already outside. */
export function visibleSelection(
  doc: ProseMirrorNode,
  selection: Selection,
  view: RevisionView,
  dir: 1 | -1,
): Selection | null {
  if (view === RevisionView.All || !(selection instanceof TextSelection)) return null
  const head = visiblePosition(doc, selection.head, view, dir)
  const anchor = selection.empty ? head : visiblePosition(doc, selection.anchor, view, dir)
  if (head === selection.head && anchor === selection.anchor) return null
  return TextSelection.between(doc.resolve(anchor), doc.resolve(head))
}

function buildDecorations(doc: ProseMirrorNode, view: RevisionView): DecorationSet {
  if (view === RevisionView.All) return DecorationSet.empty
  const decorations: Decoration[] = []
  doc.descendants((node, pos) => {
    if (isHiddenBlock(node, view)) {
      decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: 'revision-hidden' }))
      return false
    }
    if (!node.isTextblock) return true
    const classes: string[] = []
    // All text hidden but not the mark: an empty line remains.
    if (node.content.size > 0 && allContentHidden(node, view)) classes.push('revision-blank')
    if (view === RevisionView.Simple && hasRevision(node)) classes.push('revision-changed')
    if (classes.length > 0) {
      decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: classes.join(' ') }))
    }
    return false
  })
  return DecorationSet.create(doc, decorations)
}

function hasRevision(block: ProseMirrorNode): boolean {
  if (blockRevisionOf(block.attrs['markRevision']) !== null) return true
  let found = false
  block.forEach((child) => {
    if (child.marks.some((mark) => mark.type.name === INSERTION || mark.type.name === DELETION)) found = true
  })
  return found
}

export const RevisionViewExtension = Extension.create({
  name: 'revisionView',
  // Before `track-input.ts`, which would erase hidden text; after `zero-width.ts` (1000).
  priority: 900,

  addProseMirrorPlugins() {
    return [
      new Plugin<ViewState>({
        key: revisionViewKey,
        state: {
          init: () => ({ view: RevisionView.All, decorations: DecorationSet.empty }),
          apply(tr, value, _old, state) {
            const chosen = tr.getMeta(revisionViewKey) as RevisionView | undefined
            if (chosen !== undefined && chosen !== value.view) {
              return { view: chosen, decorations: buildDecorations(state.doc, chosen) }
            }
            if (!tr.docChanged || value.view === RevisionView.All) return value
            return { view: value.view, decorations: buildDecorations(state.doc, value.view) }
          },
        },
        props: {
          attributes: (state) => ({ class: `revisions-${revisionViewOf(state)}` }),
          decorations: (state) => revisionViewKey.getState(state)?.decorations ?? null,
          handleKeyDown: deleteVisibleKey,
          handleDOMEvents: { beforeinput: typeOverHidden },
          handleTextInput: replaceVisible,
        },
        appendTransaction: visibleSelectionAfter,
      }),
    ]
  },
})

/**
 * With something hidden, the browser would take the deleted text along or bring it back as plain
 * strikethrough: only what is visible is deleted here.
 */
function deleteVisibleKey(view: EditorView, event: KeyboardEvent): boolean {
  if (!isPlainDeleteKey(view, event)) return false
  const mode = revisionViewOf(view.state)
  const { selection } = view.state
  if (mode === RevisionView.All || !(selection instanceof TextSelection)) return false
  if (!selection.empty) {
    view.dispatch(deleteVisible(view.state.tr, selection.from, selection.to, mode).scrollIntoView())
    return true
  }
  return deleteVisibleCharacter(view, selection, mode, event.key === 'Backspace')
}

function isPlainDeleteKey(view: EditorView, event: KeyboardEvent): boolean {
  if (event.key !== 'Backspace' && event.key !== 'Delete') return false
  return !(event.ctrlKey || event.metaKey || event.altKey || event.shiftKey || view.composing)
}

function deleteVisibleCharacter(
  view: EditorView,
  selection: TextSelection,
  mode: RevisionView,
  backward: boolean,
): boolean {
  const $cursor = selection.$head
  const at = $cursor.start() + visibleOffset($cursor, mode, backward)
  const $at = view.state.doc.resolve(at)
  const node = backward ? $at.nodeBefore : $at.nodeAfter
  if (node === null || !node.isText || node.text === undefined) {
    // At the block edge, paragraph joining takes the usual path.
    if (at !== $cursor.pos)
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)))
    return false
  }
  const size = characterSize(node.text, backward)
  const from = backward ? at - size : at
  view.dispatch(view.state.tr.delete(from, from + size).scrollIntoView())
  return true
}

/** The cursor after skipping, in the key's direction, hidden text and zero-width marks. */
function visibleOffset($cursor: ResolvedPos, mode: RevisionView, backward: boolean): number {
  let offset = $cursor.parentOffset
  for (let moved = true; moved;) {
    moved = false
    for (const run of hiddenRuns($cursor.parent, mode)) {
      if (backward ? run.to === offset : run.from === offset) {
        offset = backward ? run.from : run.to
        moved = true
      }
    }
    const index = backward ? offset - 1 : offset
    const child = index >= 0 && index < $cursor.parent.content.size ? $cursor.parent.childAfter(index) : null
    if (child?.node !== null && child?.node !== undefined && ZERO_WIDTH.has(child.node.type.name)) {
      offset += backward ? -1 : 1
      moved = true
    }
  }
  return offset
}

/**
 * Typing over a selection with hidden text: the browser would delete everything before
 * `handleTextInput`, so the key stops at `beforeinput`.
 */
function typeOverHidden(view: EditorView, event: Event): boolean {
  const input = event as InputEvent
  const mode = revisionViewOf(view.state)
  const { selection } = view.state
  if (mode === RevisionView.All || selection.empty || view.composing) return false
  if (input.inputType !== 'insertText' || input.data === null) return false
  event.preventDefault()
  const tr = deleteVisible(view.state.tr, selection.from, selection.to, mode)
  tr.insertText(input.data, tr.mapping.map(selection.from, -1))
  view.dispatch(tr.scrollIntoView())
  return true
}

function replaceVisible(view: EditorView, from: number, to: number, text: string): boolean {
  const mode = revisionViewOf(view.state)
  if (mode === RevisionView.All || from === to || view.composing) return false
  const tr = deleteVisible(view.state.tr, from, to, mode)
  const at = tr.mapping.map(from, -1)
  tr.insertText(text, at)
  view.dispatch(tr.scrollIntoView())
  return true
}

/** A selection that landed in hidden text moves to the edge, in the direction it was going. */
function visibleSelectionAfter(
  transactions: readonly Transaction[],
  oldState: EditorState,
  newState: EditorState,
): Transaction | null {
  const mode = revisionViewOf(newState)
  if (mode === RevisionView.All) return null
  if (
    !transactions.some((tr) => tr.selectionSet || tr.docChanged || tr.getMeta(revisionViewKey) !== undefined)
  )
    return null
  const dir = newState.selection.head >= oldState.selection.head ? 1 : -1
  const selection = visibleSelection(newState.doc, newState.selection, mode, dir)
  return selection === null ? null : newState.tr.setSelection(selection)
}
