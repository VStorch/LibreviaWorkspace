import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode, ResolvedPos } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'

/**
 * Bookmark and comment ends: atomic, empty and zero-width on screen, but a position for
 * ProseMirror. What treats them as "not text" lives here.
 */
const ANCHORS = new Set(['bookmarkStart', 'bookmarkEnd', 'commentStart', 'commentEnd'])

export function isZeroWidthAnchor(node: ProseMirrorNode | null | undefined): boolean {
  return node !== null && node !== undefined && ANCHORS.has(node.type.name)
}

/** The start of the block text, after the anchors that open it. */
export function textStartOf(doc: ProseMirrorNode, pos: number): number {
  const block = doc.nodeAt(pos)
  let start = pos + 1
  for (let index = 0; block !== null && index < block.childCount; index++) {
    const child = block.child(index)
    if (!isZeroWidthAnchor(child)) break
    start += child.nodeSize
  }
  return start
}

/** The `keyCode` of keys the input method is still composing. */
const IME_PROCESS_KEY_CODE = 229

/**
 * A selection placed between anchors on purpose (go to comment): typing replaces the text and does
 * not take the comment along.
 */
export const KEEP_SELECTION = 'zeroWidthKeepSelection'

/** The selection takes the anchors touching the text edges. */
export function extendOverAnchors(state: EditorState): Transaction | null {
  const { selection } = state
  if (!(selection instanceof TextSelection) || selection.empty) return null
  const { $from, $to } = selection

  const offset = offsetBeforeAnchors($from)
  const end = offsetAfterAnchors($to)

  // Only a paragraph the selection fully covers takes the anchors.
  const sameBlock = $from.sameParent($to)
  const startCovered = offset === 0 && (!sameBlock || end === $to.parent.content.size)
  const endCovered = end === $to.parent.content.size && (!sameBlock || offset === 0)
  const from = startCovered ? $from.start() : selection.from
  const to = endCovered ? $to.end() : selection.to

  if (from === selection.from && to === selection.to) return null
  const [anchor, head] = selection.anchor <= selection.head ? [from, to] : [to, from]
  return state.tr.setSelection(TextSelection.create(state.doc, anchor, head))
}

function offsetBeforeAnchors($from: ResolvedPos): number {
  let offset = $from.parentOffset
  let before = $from.parent.childBefore(offset)
  while (offset > 0 && isZeroWidthAnchor(before.node)) {
    offset = before.offset
    before = $from.parent.childBefore(offset)
  }
  return offset
}

function offsetAfterAnchors($to: ResolvedPos): number {
  let end = $to.parentOffset
  let after = $to.parent.childAfter(end)
  while (end < $to.parent.content.size && isZeroWidthAnchor(after.node)) {
    end = after.offset + after.node!.nodeSize
    after = $to.parent.childAfter(end)
  }
  return end
}

/** Backspace and Delete skip over anchors; `null` without an anchor there. */
export function pastAnchors(state: EditorState, direction: -1 | 1): number | null {
  const { selection } = state
  if (!(selection instanceof TextSelection) || !selection.empty) return null
  return headPastAnchors(state, direction)
}

function headPastAnchors(state: EditorState, direction: -1 | 1): number | null {
  const { selection } = state
  if (!(selection instanceof TextSelection)) return null
  const $head = selection.$head
  let pos = $head.pos
  const limit = direction < 0 ? $head.start() : $head.end()
  while (pos !== limit) {
    const $pos = state.doc.resolve(pos)
    if (!isZeroWidthAnchor(direction < 0 ? $pos.nodeBefore : $pos.nodeAfter)) break
    pos += direction
  }
  return pos === $head.pos ? null : pos
}

/** The cursor or the selection head skips anchors before the browser acts; the key goes on. */
function moveOverAnchors(view: EditorView, event: KeyboardEvent): void {
  if (event.shiftKey && (event.key === 'ArrowLeft' || event.key === 'ArrowRight')) {
    // The head skips anchors before the browser extends it.
    const head = headPastAnchors(view.state, event.key === 'ArrowLeft' ? -1 : 1)
    if (head === null) return
    const anchor = view.state.selection.empty ? head : view.state.selection.anchor
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, anchor, head)))
    return
  }
  if (event.key === 'Enter') {
    // Enter with only anchors before the cursor: they follow the text.
    const back = pastAnchors(view.state, -1)
    if (back !== null && back === view.state.selection.$head.start()) moveCursor(view, back)
    return
  }
  if (event.key !== 'Backspace' && event.key !== 'Delete') return
  // Only the cursor moves; the rest continues from the new place.
  const pos = pastAnchors(view.state, event.key === 'Backspace' ? -1 : 1)
  if (pos !== null) moveCursor(view, pos)
}

function moveCursor(view: EditorView, pos: number): void {
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)))
}

/**
 * Chrome fires `selectionchange` in a separate task, and the next key may arrive first: Backspace
 * would act on the old state and delete the anchor, and the comment with it. `flush` does what
 * `selectionchange` would.
 */
export function readPendingSelection(view: EditorView): void {
  const observer = (view as unknown as { domObserver?: { flush?: () => void } }).domObserver
  observer?.flush?.()
}

/**
 * Backspace and Delete skip over anchors, which are invisible: deleting them would be a wasted key
 * press that still takes the comment. A selection reaching the edge of a paragraph's text takes the
 * anchors touching it.
 */
export const ZeroWidthAnchors = Extension.create({
  name: 'zeroWidthAnchors',
  // Before Tiptap's keymap, which would handle Backspace with the old state.
  priority: 1000,

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('zeroWidthAnchors'),
        /**
         * The browser puts the cursor inside anchors: `Shift+Home` on a caption would leave the
         * `_Ref` behind. Extended, the selection takes the anchor with the text, as in Word.
         */
        appendTransaction: (transactions, _old, state) =>
          transactions.some((tr) => tr.selectionSet) &&
          !transactions.some((tr) => tr.getMeta(KEEP_SELECTION) === true)
            ? extendOverAnchors(state)
            : null,

        props: {
          handleKeyDown(view, event) {
            if (event.isComposing || event.keyCode === IME_PROCESS_KEY_CODE) return false
            readPendingSelection(view)
            moveOverAnchors(view, event)
            return false
          },
        },
      }),
    ]
  },
})
