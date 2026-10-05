import type { Editor } from '@tiptap/react'
import type { Node as ProseMirrorNode, Schema } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { pageLabel } from '@services/document/band.js'
import { hiddenBookmarkName, nextBookmarkId } from '@services/document/bookmarks.js'
import {
  fieldArgument,
  fieldKind,
  fieldSwitch,
  sequenceNumbers,
  tocLevels,
  tocLinks,
  tocOmitsPages,
} from '@services/document/fields.js'
import type { PageSetup } from '@services/document/model.js'
import type { NoteKind } from '@services/document/notes.js'
import { outlineOf } from '@services/document/outline.js'
import {
  ensureCaptionStyle,
  ensureTocHeadingStyle,
  ensureTocStyle,
} from '@services/document/reference-styles.js'
import type { StyleSheet } from '@services/document/styles.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { bookmarksOf } from './extensions/bookmark.js'
import { noteLabelsOf, noteRefsOf, textBetweenWithoutNotes } from './extensions/note-ref.js'
import { DEFAULT_TOC_INSTRUCTION } from './extensions/table-of-contents.js'
import { readPendingSelection, textStartOf } from './extensions/zero-width.js'
import { outlineBlocksOf } from './outline-blocks.js'
import { drawnSheet, type PageLayout, type PageStart } from './usePagination.js'

/**
 * Read when the command runs: pagination changes with every line, and a stale page number is the
 * mistake a table of contents cannot have.
 */
export interface ReferenceContext {
  readonly layout: PageLayout
  readonly page: PageSetup
  /** A sheet number comes out in its section's format. When absent, `page` applies. */
  readonly sections?: readonly PageSetup[]
  readonly styles: StyleSheet
  readonly setStyles: (styles: StyleSheet) => void
  readonly t: (key: MessageKey) => string
  /** Bookmarks that exist in the file outside the nodes; see `DocumentModel.outsideBookmarks`. */
  readonly outsideBookmarks?: readonly string[]
}

/** Makes ProseMirror read now the selection the browser already changed. */
export function flushSelection(editor: Editor): void {
  readPendingSelection(editor.view)
}

/** The document position where the sheet starts, or null if the layout is stale. */
function positionOfStart(doc: ProseMirrorNode, start: PageStart): number | null {
  if (start.blockIndex >= doc.childCount) return null

  let pos = 0
  for (let index = 0; index < start.blockIndex; index++) pos += doc.child(index).nodeSize
  const block = doc.child(start.blockIndex)

  if (start.offset !== undefined) return pos + 1 + start.offset
  if (start.childIndex !== undefined) {
    let inner = pos + 1
    for (let index = 0; index < start.childIndex && index < block.childCount; index++) {
      inner += block.child(index).nodeSize
    }
    return inner
  }
  return pos
}

/** With the restart and format of the section the sheet falls in, as the `PAGE` field writes it. */
export function sheetLabel(context: ReferenceContext, sheet: number): string {
  const plan = context.layout.sheets[drawnSheet(context.layout, sheet - 1)]
  if (plan === undefined) return pageLabel(context.page, sheet)
  const section = context.sections?.[plan.section] ?? context.page
  return pageLabel({ ...section, pageNumberStart: plan.number }, 1)
}

/** The sheet (from 1) the position falls on, by the pagination breaks. */
export function sheetAt(doc: ProseMirrorNode, starts: readonly PageStart[], pos: number): number {
  let sheet = 1
  for (const start of starts) {
    const at = positionOfStart(doc, start)
    if (at === null || at > pos) break
    sheet += 1
  }
  return sheet
}

/** With each field's result in its place. */
function textBetween(doc: ProseMirrorNode, from: number, to: number): string {
  return textBetweenWithoutNotes(doc, from, to, ' ', (leaf) =>
    leaf.type.name === 'field' ? String(leaf.attrs['result'] ?? '') : '',
  )
}

/** What `NOTEREF` shows; `null` without a note in the range, and the field stays as it is. */
export function noteNumberIn(
  doc: ProseMirrorNode,
  labels: readonly string[],
  from: number,
  to: number,
): string | null {
  const index = noteRefsOf(doc).findIndex(({ pos }) => pos >= from && pos < to)
  return index < 0 ? null : (labels[index] ?? null)
}

/** Field kinds that depend on where the text falls on the sheet. */
const PAGE_KINDS = new Set(['PAGE', 'PAGEREF', 'NUMPAGES'])

/** Field kinds the editor can recompute. The rest stays as Word left it. */
const UPDATABLE = new Set(['PAGE', 'PAGEREF', 'NUMPAGES', 'REF', 'SEQ', 'NOTEREF'])

export interface FieldUpdate {
  readonly changed: number
  /** Some page field was recomputed, and pagination may change with it. */
  readonly pageDependent: boolean
}

/**
 * Word's F9, in its order: `SEQ` first, because a `REF` to a caption cites the number; then
 * references; pages last. `SEQ` counts the whole document, but only those inside the range change.
 * A bookmark that no longer exists gives "Erro! Indicador não definido.", as in Word.
 */
export function updateFieldsIn(
  editor: Editor,
  context: ReferenceContext,
  from: number,
  to: number,
  kinds: ReadonlySet<string> = UPDATABLE,
): FieldUpdate {
  // A cursor just moved may exist only in the DOM: without this the transaction would pull it back.
  flushSelection(editor)
  const { state } = editor
  const fields = fieldsOf(state.doc)
  const scan = scanFields(state, fields, context)

  const tr = state.tr
  let changed = 0
  let pageDependent = false

  for (const field of fields) {
    if (field.pos < from || field.pos >= to) continue
    if (!kinds.has(field.kind)) continue

    const update = resultOf(field, scan)
    if (update.pageDependent) pageDependent = true
    if (update.result === null || update.result === field.node.attrs['result']) continue
    tr.setNodeAttribute(field.pos, 'result', update.result)
    changed += 1
  }

  if (changed > 0) editor.view.dispatch(tr)
  return { changed, pageDependent }
}

interface FieldAt {
  readonly pos: number
  readonly node: ProseMirrorNode
  readonly kind: string
}

interface FieldScan {
  readonly doc: ProseMirrorNode
  /** The document with sequences already renumbered, which references read text from. */
  readonly sequenced: ProseMirrorNode
  readonly sequenceResult: ReadonlyMap<number, string>
  readonly bookmarks: ReadonlyMap<string, ReturnType<typeof bookmarksOf>[number]>
  readonly missing: string
  /** A bookmark outside the nodes is not lost: the result Word computed stays. */
  readonly outside: ReadonlySet<string>
  readonly noteLabels: readonly string[]
  readonly context: ReferenceContext
}

interface FieldResult {
  /** `null` leaves the field as it is. */
  readonly result: string | null
  readonly pageDependent: boolean
}

function fieldsOf(doc: ProseMirrorNode): FieldAt[] {
  const fields: FieldAt[] = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'field') fields.push({ pos, node, kind: fieldKind(String(node.attrs['instr'])) })
    return true
  })
  return fields
}

function scanFields(
  state: Editor['state'],
  fields: readonly FieldAt[],
  context: ReferenceContext,
): FieldScan {
  const sequences = fields.filter((field) => field.kind === 'SEQ')
  const numbers = sequenceNumbers(sequences.map((field) => String(field.node.attrs['instr'])))
  const sequenceResult = new Map(sequences.map((field, index) => [field.pos, numbers[index]!]))

  // A reference reads the bookmark **after** the sequences: "Figura 1" that became "Figura 2".
  const sequencing = state.tr
  for (const [pos, result] of sequenceResult) sequencing.setNodeAttribute(pos, 'result', result)
  const sequenced = sequencing.doc

  return {
    doc: state.doc,
    sequenced,
    sequenceResult,
    bookmarks: new Map(bookmarksOf(sequenced).map((bookmark) => [bookmark.name, bookmark])),
    missing: context.t('references.field.missingBookmark'),
    outside: new Set(context.outsideBookmarks ?? []),
    noteLabels: noteLabelsOf(state),
    context,
  }
}

const unpaged = (result: string | null): FieldResult => ({ result, pageDependent: false })

function resultOf(field: FieldAt, scan: FieldScan): FieldResult {
  const instr = String(field.node.attrs['instr'])
  const { context } = scan
  switch (field.kind) {
    case 'SEQ':
      return unpaged(fieldSwitch(instr, 'h') !== null ? '' : (scan.sequenceResult.get(field.pos) ?? null))
    case 'REF':
    case 'NOTEREF':
    case 'PAGEREF':
      return referenceResult(field.kind, instr, scan)
    case 'PAGE':
      return {
        result: sheetLabel(context, sheetAt(scan.doc, context.layout.pageStarts, field.pos)),
        pageDependent: true,
      }
    case 'NUMPAGES':
      return { result: String(context.layout.pages), pageDependent: true }
    default:
      return unpaged(null)
  }
}

function referenceResult(kind: 'REF' | 'NOTEREF' | 'PAGEREF', instr: string, scan: FieldScan): FieldResult {
  const name = fieldArgument(instr) ?? ''
  const target = scan.bookmarks.get(name)
  if (target === undefined) return unpaged(scan.outside.has(name) ? null : scan.missing)

  if (kind === 'REF') {
    const text = textBetween(scan.sequenced, target.pos + 1, target.end ?? target.pos + 1)
    // `\# 0`: only the number of the cited text, Word's "number only" reference.
    return unpaged(fieldSwitch(instr, '#') === null ? text : (/(\d+)(?!.*\d)/.exec(text)?.[1] ?? text))
  }
  if (kind === 'NOTEREF') {
    // The screen number, with per-sheet and per-section restarts.
    return unpaged(noteNumberIn(scan.doc, scan.noteLabels, target.pos, target.end ?? target.pos))
  }
  const sheet = sheetAt(scan.doc, scan.context.layout.pageStarts, target.pos)
  return { result: sheetLabel(scan.context, sheet), pageDependent: true }
}

/**
 * If some field depends on the page, there is a second pass once pagination settles: the new text
 * may push a line to the next sheet. Word does both too.
 */
export function updateFields(editor: Editor, context: ReferenceContext): FieldUpdate {
  const { from, to, empty } = editor.state.selection
  const range = empty ? { from: 0, to: editor.state.doc.content.size } : { from, to }
  const update = updateFieldsIn(editor, context, range.from, range.to)
  // Only when something changed: otherwise the pass would rewrite fields on the next keystroke.
  if (update.pageDependent && update.changed > 0 && empty) arm(editor, 'all')
  return update
}

/** The document's F9 fixes every page field; the table of contents, only its own. */
const pendingPagePass = new WeakMap<Editor, { scope: 'all' | 'toc'; doc: ProseMirrorNode }>()

function arm(editor: Editor, scope: 'all' | 'toc'): void {
  pendingPagePass.set(editor, { scope, doc: editor.state.doc })
}

/** Once: the pass does not request another. */
export function settlePageFields(editor: Editor, context: ReferenceContext): void {
  const pending = pendingPagePass.get(editor)
  if (pending === undefined) return
  pendingPagePass.delete(editor)
  // The user went back to typing: the pass belongs to another document.
  if (pending.doc !== editor.state.doc) return
  const { scope } = pending

  if (scope === 'all') {
    updateFieldsIn(editor, context, 0, editor.state.doc.content.size, PAGE_KINDS)
    return
  }

  for (const { pos, node } of tablesOfContents(editor.state.doc)) {
    updateFieldsIn(editor, context, pos, pos + node.nodeSize, PAGE_KINDS)
  }
}

/** Back to front, so one insertion does not shift the remaining ones. */
function ensureBookmarks(tr: Transaction, positions: readonly number[], prefix: '_Toc' | '_Ref'): string[] {
  const schema: Schema = tr.doc.type.schema
  const names = new Array<string>(positions.length)
  const used = bookmarksOf(tr.doc).map((bookmark) => bookmark.name)
  const ids: string[] = []
  tr.doc.descendants((node) => {
    if (node.type.name === 'bookmarkStart' || node.type.name === 'bookmarkEnd')
      ids.push(String(node.attrs['bid']))
    return true
  })

  const order = positions.map((pos, index) => ({ pos, index })).sort((left, right) => right.pos - left.pos)
  for (const { pos, index } of order) {
    const block = tr.doc.nodeAt(pos)
    if (block === null) continue

    let found: string | null = null
    block.forEach((child) => {
      const name = String(child.attrs['name'] ?? '')
      if (found === null && child.type.name === 'bookmarkStart' && name.startsWith(prefix)) found = name
    })
    if (found !== null) {
      names[index] = found
      continue
    }

    const name = hiddenBookmarkName(prefix, used)
    const bid = nextBookmarkId(ids)
    used.push(name)
    ids.push(bid)
    tr.insert(pos + block.nodeSize - 1, schema.nodes['bookmarkEnd']!.create({ bid }))
    tr.insert(pos + 1, schema.nodes['bookmarkStart']!.create({ name, bid }))
    names[index] = name
  }

  return names
}

function tablesOfContents(doc: ProseMirrorNode): Array<{ pos: number; node: ProseMirrorNode }> {
  const found: Array<{ pos: number; node: ProseMirrorNode }> = []
  doc.forEach((node, pos) => {
    if (node.type.name === 'tableOfContents') found.push({ pos, node })
  })
  return found
}

/**
 * As in Word: the heading text, a tab and the `PAGEREF`, inside the link to the `_Toc…` bookmark.
 */
function buildEntries(
  tr: Transaction,
  context: ReferenceContext,
  instr: string,
  sheet: StyleSheet,
): { entries: unknown[]; sheet: StyleSheet } {
  const { from, to } = tocLevels(instr)
  const links = tocLinks(instr)
  const omitPages = tocOmitsPages(instr)

  // An old table of contents with paragraphs in heading styles would list itself.
  const inside = tablesOfContents(tr.doc).map(({ pos, node }) => [pos, pos + node.nodeSize] as const)
  const headings = outlineOf(outlineBlocksOf(tr.doc), sheet).filter(
    (heading) =>
      heading.level >= from &&
      heading.level <= to &&
      !inside.some(([start, end]) => heading.pos > start && heading.pos < end),
  )

  const names = ensureBookmarks(
    tr,
    headings.map((heading) => heading.pos),
    '_Toc',
  )

  let styles = sheet
  const entries: unknown[] = headings.map((heading, index) => {
    const ensured = ensureTocStyle(styles, heading.level)
    styles = ensured.sheet
    const name = names[index]!
    const marks = links ? [{ type: 'link', attrs: { href: `#${name}` } }] : []
    // The current pagination's sheet, which measured the document before the bookmarks; the second
    // pass corrects it.
    const page = sheetLabel(context, sheetAt(tr.before, context.layout.pageStarts, heading.pos))
    return {
      type: 'paragraph',
      attrs: { styleId: ensured.id },
      content: [
        { type: 'text', text: omitPages ? heading.text : `${heading.text}\t`, marks },
        ...(omitPages
          ? []
          : [{ type: 'field', attrs: { instr: ` PAGEREF ${name} \\h `, result: page }, marks }]),
      ],
    }
  })

  if (entries.length === 0) {
    entries.push({
      type: 'paragraph',
      attrs: { styleId: null },
      content: [{ type: 'text', text: context.t('references.toc.empty') }],
    })
  }

  return { entries, sheet: styles }
}

/** Before the cursor's block, or in its place if empty, with headings 1 to 3. */
export function insertTableOfContents(editor: Editor, context: ReferenceContext): void {
  const { state } = editor
  const tr = state.tr

  const heading = ensureTocHeadingStyle(context.styles)
  const { entries, sheet } = buildEntries(tr, context, DEFAULT_TOC_INSTRUCTION, heading.sheet)

  const title = {
    type: 'paragraph',
    attrs: { styleId: heading.id },
    content: [{ type: 'text', text: context.t('references.toc.title') }],
  }
  const node = state.schema.nodeFromJSON({
    type: 'tableOfContents',
    attrs: { instr: DEFAULT_TOC_INSTRUCTION, head: 1, sdt: true },
    content: [title, ...entries],
  })

  const $from = tr.doc.resolve(tr.mapping.map(state.selection.from))
  const top = $from.depth === 0 ? $from.pos : $from.before(1)
  const current = tr.doc.nodeAt(top)
  if (current !== null && current.isTextblock && current.content.size === 0) {
    tr.replaceWith(top, top + current.nodeSize, node)
  } else {
    tr.insert(top, node)
  }

  if (sheet !== context.styles) context.setStyles(sheet)
  editor.view.dispatch(tr.scrollIntoView())
  arm(editor, 'toc')
}

/**
 * The cursor's, or the first. The title, the instruction and the content control stay; the entries
 * are replaced whole, as in Word's "Update entire table". False when there is no table of contents.
 */
export function updateTableOfContents(editor: Editor, context: ReferenceContext): boolean {
  const { state } = editor
  const all = tablesOfContents(state.doc)
  if (all.length === 0) return false

  const cursor = state.selection.from
  const chosen = all.find(({ pos, node }) => cursor > pos && cursor < pos + node.nodeSize) ?? all[0]!
  const tr = state.tr
  const instr = String(chosen.node.attrs['instr'] ?? DEFAULT_TOC_INSTRUCTION)
  const { entries, sheet } = buildEntries(tr, context, instr, context.styles)

  // The new bookmarks may have come in before the table of contents.
  const pos = tr.mapping.map(chosen.pos)
  const toc = tr.doc.nodeAt(pos)
  if (toc === null) return false

  const head = Math.min(Number(toc.attrs['head'] ?? 0), toc.childCount - 1)
  const kept: ProseMirrorNode[] = []
  for (let index = 0; index < head; index++) kept.push(toc.child(index))
  const fresh = entries.map((entry) => state.schema.nodeFromJSON(entry))

  tr.replaceWith(pos + 1, pos + toc.nodeSize - 1, [...kept, ...fresh])

  if (sheet !== context.styles) context.setStyles(sheet)
  editor.view.dispatch(tr)
  arm(editor, 'toc')
  return true
}

function sequencesIn(block: ProseMirrorNode, pos: number): Array<{ pos: number; label: string }> {
  const found: Array<{ pos: number; label: string }> = []
  block.forEach((child, offset) => {
    if (child.type.name !== 'field') return
    const instr = String(child.attrs['instr'] ?? '')
    if (fieldKind(instr) === 'SEQ') found.push({ pos: pos + 1 + offset, label: fieldArgument(instr) ?? '' })
  })
  return found
}

/** The ones the document already uses and Word's (Figure, Table, Equation) in the UI language. */
export function captionLabels(doc: ProseMirrorNode, defaults: readonly string[]): string[] {
  const labels = new Set(defaults)
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    for (const sequence of sequencesIn(node, pos)) if (sequence.label !== '') labels.add(sequence.label)
    return false
  })
  return [...labels]
}

export interface CaptionRequest {
  readonly label: string
  /** May be empty. */
  readonly text: string
  /** Above is the custom for tables; below, for figures. */
  readonly above: boolean
}

/**
 * In the `caption` style, with a `SEQ` already counted, as in Word. Later captions are only
 * renumbered by "Update fields", also as in Word.
 */
export function insertCaption(editor: Editor, context: ReferenceContext, request: CaptionRequest): void {
  flushSelection(editor)
  const { state } = editor
  const ensured = ensureCaptionStyle(context.styles)
  const label = request.label.trim()
  if (label === '') return

  const $from = state.selection.$from
  const top = $from.depth === 0 ? $from.pos : $from.before(1)
  const block = state.doc.nodeAt(top)
  const at = request.above || block === null ? top : top + block.nodeSize

  const text = request.text.trim()
  const paragraph = state.schema.nodeFromJSON({
    type: 'paragraph',
    attrs: { styleId: ensured.id },
    content: [
      { type: 'text', text: `${label} ` },
      { type: 'field', attrs: { instr: ` SEQ ${label} \\* ARABIC `, result: '1' } },
      ...(text === '' ? [] : [{ type: 'text', text: ` ${text}` }]),
    ],
  })

  const tr = state.tr.insert(at, paragraph)
  if (ensured.sheet !== context.styles) context.setStyles(ensured.sheet)
  editor.view.dispatch(tr.scrollIntoView())

  const fieldPos = at + 1 + label.length + 1
  updateFieldsIn(editor, context, fieldPos, fieldPos + 1, new Set(['SEQ']))
}

/** What the reference points to: a heading, a bookmark, or a label's caption. */
export type CrossReferenceKind =
  | { readonly type: 'heading' }
  | { readonly type: 'bookmark' }
  | { readonly type: 'caption'; readonly label: string }
  | { readonly type: 'note'; readonly kind: NoteKind }

export interface CrossReferenceTarget {
  readonly key: string
  readonly text: string
}

export function crossReferenceTargets(
  doc: ProseMirrorNode,
  sheet: StyleSheet,
  kind: CrossReferenceKind,
  /** To list notes by their screen number (`noteLabelsOf`). */
  labels?: readonly string[],
): CrossReferenceTarget[] {
  if (kind.type === 'heading') {
    return outlineOf(outlineBlocksOf(doc), sheet).map((entry) => ({
      key: String(entry.pos),
      text: `${' '.repeat(entry.level - 1)}${entry.text}`,
    }))
  }

  if (kind.type === 'bookmark') {
    return bookmarksOf(doc)
      .filter((bookmark) => !bookmark.name.startsWith('_'))
      .map((bookmark) => ({ key: bookmark.name, text: bookmark.name }))
  }

  if (kind.type === 'note') {
    // The number and the start of the text, as Word lists them.
    return noteRefsOf(doc).flatMap(({ node, pos }, index) => {
      if (node.attrs['kind'] !== kind.kind) return []
      const text = node.textBetween(0, node.content.size, ' ').trim()
      const label = labels?.[index] ?? ''
      const short = text.length > 60 ? `${text.slice(0, 60)}…` : text
      return [{ key: String(pos), text: `${label} ${short}`.trim() }]
    })
  }

  const wanted = kind.label.toLowerCase()
  const captions: CrossReferenceTarget[] = []
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    if (sequencesIn(node, pos).some((sequence) => sequence.label.toLowerCase() === wanted)) {
      captions.push({ key: String(pos), text: textBetween(doc, pos + 1, pos + node.nodeSize - 1).trim() })
    }
    return false
  })
  return captions
}

export type CrossReferenceShow = 'text' | 'number' | 'page'

export interface CrossReferenceRequest {
  readonly kind: CrossReferenceKind
  readonly key: string
  readonly show: CrossReferenceShow
  /** `\h`: with Ctrl+click, the reference leads to the target. */
  readonly link: boolean
}

/**
 * An existing one counts when it ends right after the range and starts up to `slack` positions
 * before: Word's caption bookmark starts before "Figura".
 */
function rangeBookmark(tr: Transaction, from: number, to: number, slack: number): string {
  const existing = bookmarksOf(tr.doc).find(
    (bookmark) =>
      bookmark.name.startsWith('_Ref') &&
      bookmark.end === to &&
      bookmark.pos < from &&
      bookmark.pos >= from - 1 - slack,
  )
  if (existing !== undefined) return existing.name

  const ids: string[] = []
  tr.doc.descendants((node) => {
    if (node.type.name === 'bookmarkStart' || node.type.name === 'bookmarkEnd')
      ids.push(String(node.attrs['bid']))
    return true
  })
  const name = hiddenBookmarkName(
    '_Ref',
    bookmarksOf(tr.doc).map((bookmark) => bookmark.name),
  )
  const bid = nextBookmarkId(ids)
  const schema = tr.doc.type.schema
  tr.insert(to, schema.nodes['bookmarkEnd']!.create({ bid }))
  tr.insert(from, schema.nodes['bookmarkStart']!.create({ name, bid }))
  return name
}

/**
 * Headings and captions get the hidden bookmark right away, as in Word. For a caption, "text" is
 * label and number, and "number" only the number.
 */
export function insertCrossReference(
  editor: Editor,
  context: ReferenceContext,
  request: CrossReferenceRequest,
): boolean {
  flushSelection(editor)
  const { state } = editor
  const tr = state.tr
  let name: string | null

  if (request.kind.type === 'bookmark') {
    name = request.key
  } else if (request.kind.type === 'heading') {
    name = ensureBookmarks(tr, [Number(request.key)], '_Ref')[0] ?? null
  } else if (request.kind.type === 'note') {
    // The bookmark `NOTEREF` cites, as Word writes it.
    const pos = Number(request.key)
    const reference = tr.doc.nodeAt(pos)
    if (reference === null || reference.type.name !== 'noteRef') return false
    name = rangeBookmark(tr, pos, pos + reference.nodeSize, 0)
  } else {
    const pos = Number(request.key)
    const block = tr.doc.nodeAt(pos)
    const wanted = request.kind.label.toLowerCase()
    const sequence =
      block === null ? undefined : sequencesIn(block, pos).find((item) => item.label.toLowerCase() === wanted)
    if (sequence === undefined) return false
    const first = textStartOf(tr.doc, pos)
    // One bookmark per caption, for the text, the number (`\# 0`) and the page, as in Word.
    name = rangeBookmark(tr, first, sequence.pos + 1, first - pos - 1)
  }
  if (name === null) return false

  const switches = `${request.show === 'number' ? ' \\# 0' : ''}${request.link ? ' \\h' : ''}`
  const instr =
    request.show === 'page'
      ? ` PAGEREF ${name}${switches} `
      : request.kind.type === 'note'
        ? ` NOTEREF ${name}${request.link ? ' \\h' : ''} `
        : ` REF ${name}${switches} `
  const at = tr.mapping.map(state.selection.from)
  tr.replaceWith(
    at,
    tr.mapping.map(state.selection.to),
    state.schema.nodes['field']!.create({ instr, result: '' }),
  )
  editor.view.dispatch(tr.scrollIntoView())

  // The same path as F9.
  updateFieldsIn(editor, context, at, at + 1)
  return true
}
