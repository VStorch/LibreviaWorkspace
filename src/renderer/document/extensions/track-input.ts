import { Extension } from '@tiptap/core'
import { Fragment, Slice, type Mark, type Node as ProseMirrorNode, type ResolvedPos } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Mapping, ReplaceStep, StepMap, canJoin, type Step } from '@tiptap/pm/transform'
import type { EditorView } from '@tiptap/pm/view'
import { DELETION, INSERTION, ZERO_WIDTH, blockRevisionOf, characterSize } from './track-changes.js'

/**
 * The transaction is rewritten **before** being applied (`dispatchTransaction`): a single one, with
 * the right selection and one undo step. In each `ReplaceStep`:
 *
 * - what comes in gets the author's `insertion` and loses `deletion`;
 * - what goes out **stays**, with `deletion`, except the author's own insertion, which really goes
 *   as in Word, and text already deleted, which stays as it was;
 * - an Enter that goes out becomes `markRevision del` on the block above, and one that comes in,
 *   `markRevision ins`;
 * - a whole table row gets `rowRevision`.
 *
 * Formatting passes untracked, as do undo, accept and reject (`SKIP_TRACKING`) and a lone comment
 * or bookmark end. During IME composition only the insertion is marked: bringing back what it
 * deletes would break its DOM.
 */

/** Accept and reject, for instance. */
export const SKIP_TRACKING = 'trackChanges:skip'

/** Undo and redo are not new edits. */
const HISTORY_META = 'history$'

/** To the minute, as in Word: what is typed in the same minute merges into one range. */
export function revisionDate(now: Date): string {
  return `${now.toISOString().slice(0, 16)}:00Z`
}

interface Revision {
  readonly author: string
  readonly date: string
}

export interface TrackOptions {
  readonly composing?: boolean
}

export function shouldTrack(tr: Transaction): boolean {
  if (!tr.docChanged) return false
  if (tr.getMeta(SKIP_TRACKING) === true) return false
  if (tr.getMeta(HISTORY_META) !== undefined) return false
  return tr.getMeta('addToHistory') !== false
}

/**
 * The deleted text tracking kept, seen from the original: each point is a position in the original
 * where the tracked document has `size` extra positions.
 */
class KeptContent {
  private points: { pos: number; size: number }[] = []

  /** At the exact point, `assoc` says which side. */
  map(pos: number, assoc: -1 | 1): number {
    let result = pos
    for (const point of this.points) {
      if (point.pos < pos || (point.pos === pos && assoc > 0)) result += point.size
    }
    return result
  }

  mapping(): Mapping {
    const mapping = new Mapping()
    let shift = 0
    for (const point of this.points) {
      mapping.appendMap(new StepMap([point.pos + shift, 0, point.size]))
      shift += point.size
    }
    return mapping
  }

  /** Points inside `absorbed` collapse to one at its start; those the step deleted disappear. */
  advance(map: StepMap, absorbed?: { from: number; to: number; size: number }): void {
    const next: { pos: number; size: number }[] = []
    for (const point of this.points) {
      if (absorbed !== undefined && point.pos >= absorbed.from && point.pos <= absorbed.to) continue
      const result = map.mapResult(point.pos, -1)
      if (result.deletedAcross) continue
      next.push({ pos: result.pos, size: point.size })
    }
    if (absorbed !== undefined && absorbed.size > 0) next.push({ pos: absorbed.from, size: absorbed.size })
    next.sort((a, b) => a.pos - b.pos)
    this.points = []
    for (const point of next) {
      const last = this.points[this.points.length - 1]
      if (last !== undefined && last.pos === point.pos) last.size += point.size
      else this.points.push({ ...point })
    }
  }
}

function hasMark(node: ProseMirrorNode, name: string): boolean {
  return node.marks.some((mark) => mark.type.name === name)
}

/** Deleting what the author themselves inserted is really deleting. */
function isOwnInsertion(node: ProseMirrorNode, author: string): boolean {
  return (
    !hasMark(node, DELETION) &&
    node.marks.some((mark) => mark.type.name === INSERTION && mark.attrs['author'] === author)
  )
}

function isOwnBlockInsertion(value: unknown, author: string): boolean {
  const revision = blockRevisionOf(value)
  return revision?.kind === 'ins' && revision.author === author
}

function blockRevision(kind: 'ins' | 'del', revision: Revision): Record<string, string> {
  return { kind, author: revision.author, date: revision.date }
}

interface DeletionPlan {
  readonly marks: [number, number][]
  /** Ranges that really go: own insertion, block object, own row. */
  readonly drops: [number, number][]
  /** Blocks whose paragraph mark becomes deleted. */
  readonly paragraphMarks: number[]
  /** Blocks whose paragraph mark (inserted by the author) goes: they join the next one. */
  readonly joins: number[]
  readonly rows: number[]
}

/**
 * The block's own, when its end falls in the range; the one above, on Backspace in an empty
 * paragraph.
 */
function paragraphMarkTaken(
  doc: ProseMirrorNode,
  node: ProseMirrorNode,
  pos: number,
  from: number,
  to: number,
): number | null {
  const end = pos + node.nodeSize
  if (end - 1 < from || end > to) return null
  if (end < to) return pos
  if (pos < from) return null
  const $pos = doc.resolve(pos)
  const index = $pos.index()
  if (index < $pos.parent.childCount - 1) return pos
  const previous = index > 0 ? $pos.parent.child(index - 1) : null
  if (previous?.isTextblock !== true) return null
  return pos - previous.nodeSize
}

/** `null` when the deletion is untracked, as in a table column. */
function planDeletion(doc: ProseMirrorNode, from: number, to: number, author: string): DeletionPlan | null {
  const planner = new DeletionPlanner(doc, from, to, author)
  doc.nodesBetween(from, to, (node, pos) => planner.visit(node, pos))
  return planner.untracked ? null : planner.plan
}

class DeletionPlanner {
  readonly plan: DeletionPlan = { marks: [], drops: [], paragraphMarks: [], joins: [], rows: [] }
  untracked = false
  private readonly taken = new Set<number>()

  constructor(
    private readonly doc: ProseMirrorNode,
    private readonly from: number,
    private readonly to: number,
    private readonly author: string,
  ) {}

  /** The return value is `nodesBetween`'s: whether to descend into children. */
  visit(node: ProseMirrorNode, pos: number): boolean {
    if (this.untracked) return false
    const end = pos + node.nodeSize

    if (node.isInline) {
      this.inline(node, pos, end)
      return false
    }

    if (node.isTextblock) {
      this.paragraphMark(node, pos)
      return true
    }

    return pos >= this.from && end <= this.to ? this.wholeBlock(node, pos, end) : true
  }

  private wholeBlock(node: ProseMirrorNode, pos: number, end: number): boolean {
    if (node.type.name === 'tableRow') {
      const revision = blockRevisionOf(node.attrs['rowRevision'])
      if (isOwnBlockInsertion(node.attrs['rowRevision'], this.author)) this.plan.drops.push([pos, end])
      else if (revision?.kind !== 'del') this.plan.rows.push(pos)
      return false
    }

    // A whole cell without the whole row is a column: Word does not track it this way, and a
    // document with cell revisions already opens locked.
    if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
      this.untracked = true
      return false
    }

    if (node.isBlock && node.isLeaf) {
      this.plan.drops.push([pos, end])
      return false
    }
    return true
  }

  private inline(node: ProseMirrorNode, pos: number, end: number): void {
    const start = Math.max(pos, this.from)
    const stop = Math.min(end, this.to)
    if (stop <= start || ZERO_WIDTH.has(node.type.name) || hasMark(node, DELETION)) return
    if (isOwnInsertion(node, this.author)) this.plan.drops.push([start, stop])
    else this.plan.marks.push([start, stop])
  }

  private paragraphMark(node: ProseMirrorNode, pos: number): void {
    const owner = paragraphMarkTaken(this.doc, node, pos, this.from, this.to)
    if (owner === null || this.taken.has(owner)) return
    this.taken.add(owner)
    const block = this.doc.nodeAt(owner)!
    const after = owner + block.nodeSize
    const revision = blockRevisionOf(block.attrs['markRevision'])
    if (isOwnBlockInsertion(block.attrs['markRevision'], this.author) && canJoin(this.doc, after))
      this.plan.joins.push(owner)
    else if (revision?.kind !== 'del') this.plan.paragraphMarks.push(owner)
  }
}

/** Returns how many positions really went. */
function applyDeletion(tr: Transaction, plan: DeletionPlan, revision: Revision): number {
  const schema = tr.doc.type.schema
  const sizeBefore = tr.doc.content.size
  const mark = schema.marks[DELETION]!.create({ ...revision })

  for (const pos of plan.paragraphMarks) {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, markRevision: blockRevision('del', revision) })
  }
  for (const pos of plan.rows) {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, rowRevision: blockRevision('del', revision) })
  }
  for (const [from, to] of plan.marks) tr.addMark(from, to, mark)

  // A joining block takes the next one's paragraph mark. Back to front, so what goes further ahead
  // does not move earlier positions.
  const removals: { from: number; to: number; join?: number }[] = plan.drops.map(([from, to]) => ({
    from,
    to,
  }))
  for (const pos of plan.joins) {
    const end = pos + tr.doc.nodeAt(pos)!.nodeSize
    removals.push({ from: end - 1, to: end + 1, join: pos })
  }
  removals.sort((a, b) => b.from - a.from)
  for (const { from, to, join } of removals) {
    if (join !== undefined) {
      const node = tr.doc.nodeAt(join)!
      const next = tr.doc.nodeAt(join + node.nodeSize)
      tr.setNodeMarkup(join, undefined, { ...node.attrs, markRevision: next?.attrs['markRevision'] ?? null })
    }
    tr.delete(from, to)
  }

  return sizeBefore - tr.doc.content.size
}

/**
 * The paragraph ProseMirror puts in place of the deleted one; when open, it is the Enter, and it
 * comes in.
 */
function onlyEmptyBlocks(slice: Slice): boolean {
  if (slice.openStart > 0) return false
  let empty = true
  slice.content.forEach((node) => {
    if (!node.isTextblock || node.content.size > 0) empty = false
  })
  return empty
}

function markInserted(tr: Transaction, from: number, to: number, revision: Revision): void {
  const schema = tr.doc.type.schema
  const insertion = schema.marks[INSERTION]!.create({ ...revision })
  const deletion = schema.marks[DELETION]!
  const inline: [number, number][] = []
  const paragraphs: number[] = []
  const rows: number[] = []

  tr.doc.nodesBetween(from, to, (node, pos) => {
    const end = pos + node.nodeSize
    if (node.isInline) {
      if (!ZERO_WIDTH.has(node.type.name)) inline.push([Math.max(pos, from), Math.min(end, to)])
      return false
    }
    if (node.type.name === 'tableRow' && pos >= from && end <= to) {
      rows.push(pos)
      return false
    }
    if (node.isTextblock && end - 1 >= from && end - 1 < to) paragraphs.push(pos)
    return true
  })

  for (const [start, stop] of inline) {
    if (stop <= start) continue
    tr.removeMark(start, stop, deletion)
    tr.addMark(start, stop, insertion)
  }
  for (const pos of paragraphs) {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, markRevision: blockRevision('ins', revision) })
  }
  for (const pos of rows) {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, rowRevision: blockRevision('ins', revision) })
  }
}

function onlyAnchors(step: ReplaceStep, doc: ProseMirrorNode): boolean {
  const anchorsOnly = (fragment: Fragment): boolean => {
    let only = true
    fragment.descendants((node) => {
      if (!ZERO_WIDTH.has(node.type.name)) only = false
      return false
    })
    return only
  }
  return anchorsOnly(doc.slice(step.from, step.to).content) && anchorsOnly(step.slice.content)
}

function copyExtras(from: Transaction, to: Transaction): void {
  const meta = (from as unknown as { meta: Record<string, unknown> }).meta
  for (const key of Object.keys(meta)) to.setMeta(key, meta[key])
  to.setTime(from.time)
  if (from.scrolledIntoView) to.scrollIntoView()
}

/** Returns how much the document shrank. */
function deleteForReal(tracked: Transaction, from: number, to: number): number {
  const sizeBefore = tracked.doc.content.size
  tracked.delete(from, to)
  return sizeBefore - tracked.doc.content.size
}

function insertTracked(tracked: Transaction, to: number, slice: Slice, revision: Revision): void {
  const before = tracked.steps.length
  // The tracked document's paragraph mark, not the original's.
  const $at = tracked.doc.resolve(to)
  const tail: unknown = $at.parent.isTextblock ? $at.parent.attrs['markRevision'] : undefined
  tracked.replace(to, to, slice)
  if (tracked.steps.length === before) return

  const mapping = tracked.mapping.slice(before)
  const end = mapping.map(to, 1)
  markInserted(tracked, mapping.map(to, -1), end, revision)
  const $end = tracked.doc.resolve(end)
  if (slice.openEnd > 0 && tail !== undefined && $end.parent.isTextblock && $end.depth > 0) {
    if ($end.parent.attrs['markRevision'] !== tail) {
      tracked.setNodeMarkup($end.before(), undefined, { ...$end.parent.attrs, markRevision: tail })
    }
  }
}

/** Pure: takes the previous state and returns another transaction on it. */
export function trackTransaction(
  tr: Transaction,
  state: EditorState,
  author: string,
  now: Date,
  options: TrackOptions = {},
): Transaction {
  const revision: Revision = { author, date: revisionDate(now) }
  const tracked = state.tr
  const kept = new KeptContent()

  tr.steps.forEach((step: Step, index) => {
    const doc = tr.docs[index]!
    const map = step.getMap()
    const passThrough = (): void => {
      const mapped = step.map(kept.mapping())
      if (mapped !== null) tracked.maybeStep(mapped)
      kept.advance(map)
    }

    if (!(step instanceof ReplaceStep) || onlyAnchors(step, doc)) return passThrough()

    const from = kept.map(step.from, -1)
    let to = kept.map(step.to, 1)
    const deletes = step.to > step.from

    // During composition, what goes out really goes; only what comes in is marked.
    const tracksDeletion = deletes && options.composing !== true
    const plan = tracksDeletion ? planDeletion(tracked.doc, from, to, author) : null
    if (tracksDeletion && plan === null) return passThrough()

    if (plan !== null) to -= applyDeletion(tracked, plan, revision)
    else if (deletes) to -= deleteForReal(tracked, from, to)

    // What comes in goes after the deleted text, as in Word.
    if (step.slice.size > 0 && !(deletes && onlyEmptyBlocks(step.slice))) {
      insertTracked(tracked, to, step.slice, revision)
    }

    kept.advance(map, { from: step.from, to: step.to, size: to - from })
  })

  copyExtras(tr, tracked)

  // Backspace leaves the cursor before the deleted text; Delete and the rest, after it.
  const before = state.selection
  const after = tr.selection
  const backward = before.empty && after.empty && after.head < before.head
  const assoc = backward ? -1 : 1
  if (after instanceof TextSelection) {
    const doc = tracked.doc
    const anchor = Math.min(kept.map(after.anchor, assoc), doc.content.size)
    const head = Math.min(kept.map(after.head, assoc), doc.content.size)
    tracked.setSelection(TextSelection.between(doc.resolve(anchor), doc.resolve(head)))
  } else {
    tracked.setSelection(after.map(tracked.doc, kept.mapping()))
  }
  if (tr.storedMarksSet) tracked.setStoredMarks(tr.storedMarks)

  return tracked
}

export interface TrackGroup {
  readonly id: string
  readonly time: number
  readonly head: number
}

/** `prosemirror-history`'s `newGroupDelay`. */
const GROUP_DELAY_MS = 500
let groupCount = 0

/**
 * A tracked deletion only adds a mark, and the history, which groups by the proximity of changed
 * ranges, would make each Backspace a step. Here proximity is that of the requested edit, and
 * grouping goes through the `composition` meta, the same the history uses for IME.
 */
export function joinHistoryGroup(
  original: Transaction,
  tracked: Transaction,
  previous: TrackGroup | null,
): TrackGroup | null {
  if (original.getMeta('composition') !== undefined) return null
  const map = original.mapping.maps[0]
  let adjacent = false
  if (previous !== null && original.time - previous.time < GROUP_DELAY_MS && map !== undefined) {
    map.forEach((start, end) => {
      if (start <= previous.head && end >= previous.head) adjacent = true
    })
  }
  const id = adjacent && previous !== null ? previous.id : `track-${++groupCount}`
  tracked.setMeta('composition', id)
  return { id, time: original.time, head: tracked.selection.head }
}

function isRevisionMark(mark: Mark): boolean {
  return mark.type.name === INSERTION || mark.type.name === DELETION
}

/** What is pasted is new text. */
function withoutRevisions(fragment: Fragment): Fragment {
  const children: ProseMirrorNode[] = []
  fragment.forEach((node) => {
    if (node.isText) {
      children.push(node.mark(node.marks.filter((mark) => !isRevisionMark(mark))))
      return
    }
    const attrs =
      'markRevision' in node.attrs || 'rowRevision' in node.attrs
        ? {
            ...node.attrs,
            ...('markRevision' in node.attrs ? { markRevision: null } : {}),
            ...('rowRevision' in node.attrs ? { rowRevision: null } : {}),
          }
        : node.attrs
    children.push(
      node.type.create(
        attrs,
        withoutRevisions(node.content),
        node.marks.filter((mark) => !isRevisionMark(mark)),
      ),
    )
  })
  return Fragment.fromArray(children)
}

/** Copying takes the text as it will be. */
function withoutDeleted(fragment: Fragment): Fragment {
  const children: ProseMirrorNode[] = []
  fragment.forEach((node) => {
    if (node.isInline && hasMark(node, DELETION)) return
    children.push(node.isLeaf ? node : node.copy(withoutDeleted(node.content)))
  })
  return Fragment.fromArray(children)
}

export function stripRevisions(slice: Slice): Slice {
  return new Slice(withoutRevisions(slice.content), slice.openStart, slice.openEnd)
}

export function stripDeleted(slice: Slice): Slice {
  return new Slice(withoutDeleted(slice.content), slice.openStart, slice.openEnd)
}

export interface TrackInputOptions {
  /** Queried on each transaction: it changes while the editor runs. */
  readonly isTracking: () => boolean
  readonly author: () => string
}

const trackInputKey = new PluginKey('trackInput')

function hasRevisionInside(block: ProseMirrorNode): boolean {
  let found = false
  block.forEach((child) => {
    if (child.marks.some((mark) => mark.type.name === INSERTION || mark.type.name === DELETION)) found = true
  })
  return found
}

/** Already deleted text and zero-width anchors are transparent; `null` at the block edge. */
export function wordRangeAt(
  block: ProseMirrorNode,
  offset: number,
  backward: boolean,
): [number, number] | null {
  // 'w' letter, 's' space, 't' transparent.
  const kinds: string[] = []
  block.forEach((child) => {
    if (child.isText) {
      const deleted = hasMark(child, DELETION)
      for (const char of child.text ?? '') {
        const kind = deleted ? 't' : /\s/.test(char) ? 's' : 'w'
        for (let unit = 0; unit < char.length; unit++) kinds.push(kind)
      }
    } else {
      const kind = ZERO_WIDTH.has(child.type.name) ? 't' : 'w'
      for (let unit = 0; unit < child.nodeSize; unit++) kinds.push(kind)
    }
  })

  let index = offset
  if (backward) {
    while (index > 0 && kinds[index - 1] !== 'w') index--
    while (index > 0 && kinds[index - 1] !== 's') index--
    return index === offset ? null : [index, offset]
  }
  while (index < kinds.length && kinds[index] !== 's') index++
  while (index < kinds.length && kinds[index] !== 'w') index++
  return index === offset ? null : [offset, index]
}

/**
 * A note reference only as its mark: `textBetween` would dump the body in the middle of the
 * sentence.
 */
function plainTextOf(fragment: Fragment): string {
  const blocks: string[] = []
  let inline = ''
  fragment.forEach((node) => {
    if (node.isText) inline += node.text ?? ''
    else if (node.type.name === 'hardBreak') inline += '\n'
    else if (node.isInline) inline += node.type.name === 'noteRef' ? String(node.attrs['mark'] ?? '') : ''
    else blocks.push(plainTextOf(node.content))
  })
  if (inline !== '') blocks.unshift(inline)
  return blocks.join('\n\n')
}

/**
 * Done here, not by the browser, which would touch the neighbouring deleted range and bring it back
 * as new text. Exported for the note body, which is another `EditorView` without the editor
 * plugins.
 */
export function trackedDeleteKey(view: EditorView, event: KeyboardEvent, isTracking: () => boolean): boolean {
  if (view.composing) return false
  if (event.key !== 'Backspace' && event.key !== 'Delete') return false
  if (event.metaKey || event.altKey || event.shiftKey) return false
  const { selection } = view.state
  if (!selection.empty) return false
  const backward = event.key === 'Backspace'

  // The word too: the browser would redo the neighbouring `<del>` as plain strikethrough.
  if (event.ctrlKey) return deleteWord(view, selection.$head, backward, isTracking)
  if (!isTracking()) return false
  return deleteCharacter(view, selection.$head, backward)
}

function deleteWord(
  view: EditorView,
  $cursor: ResolvedPos,
  backward: boolean,
  isTracking: () => boolean,
): boolean {
  if (!isTracking() && !hasRevisionInside($cursor.parent)) return false
  const range = wordRangeAt($cursor.parent, $cursor.parentOffset, backward)
  if (range === null) return false
  const start = $cursor.start()
  view.dispatch(view.state.tr.delete(start + range[0], start + range[1]).scrollIntoView())
  return true
}

function deleteCharacter(view: EditorView, $cursor: ResolvedPos, backward: boolean): boolean {
  const node = backward ? $cursor.nodeBefore : $cursor.nodeAfter
  if (node === null || !node.isText || node.text === undefined) return false
  const size = characterSize(node.text, backward)
  const from = backward ? $cursor.pos - size : $cursor.pos
  view.dispatch(view.state.tr.delete(from, from + size).scrollIntoView())
  return true
}

export const TrackInput = Extension.create<TrackInputOptions>({
  name: 'trackInput',

  addOptions() {
    return { isTracking: () => false, author: () => '' }
  },

  addStorage() {
    return { group: null as TrackGroup | null }
  },

  dispatchTransaction({ transaction, next }) {
    if (!this.options.isTracking() || !shouldTrack(transaction)) {
      // Another edit in between: the next tracked one starts a new group.
      if (transaction.docChanged) this.storage.group = null
      next(transaction)
      return
    }
    const view = this.editor.view
    let tracked: Transaction
    try {
      tracked = trackTransaction(transaction, view.state, this.options.author(), new Date(), {
        composing: view.composing,
      })
    } catch (error) {
      // Better an untracked edit than a lost one.
      console.error(error)
      tracked = transaction
    }
    if (tracked !== transaction)
      this.storage.group = joinHistoryGroup(transaction, tracked, this.storage.group)
    next(tracked)
  },

  addProseMirrorPlugins() {
    const options = this.options
    return [
      new Plugin({
        key: trackInputKey,
        props: {
          handleKeyDown: (view, event) => trackedDeleteKey(view, event, options.isTracking),
          // Pasted content comes in as new text; deleted text does not go to the clipboard, not
          // even as plain text.
          transformPasted: (slice) => stripRevisions(slice),
          transformCopied: (slice) => stripDeleted(slice),
          clipboardTextSerializer: (slice) => {
            return plainTextOf(stripDeleted(slice).content)
          },
        },
      }),
    ]
  },
})
