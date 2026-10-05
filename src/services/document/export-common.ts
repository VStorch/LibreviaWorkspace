import { ALLOWED_EXTERNAL_PROTOCOLS } from '@shared/constants.js'
import { numberLists, LIST_TYPES, type ListInfo, type ListTreeReader } from './list-numbering.js'
import type { DocumentModel, DocumentNode } from './model.js'
import { NoteKind, noteCounter } from './notes.js'

/**
 * From the **model**, not the editor HTML, which carries decorations and attributes only the editor
 * understands.
 */

export interface ExportNote {
  readonly kind: NoteKind
  readonly label: string
  /** The note anchor (`nota-rodape-1`, `nota-fim-1`); the reference's has `ref-` in front. */
  readonly id: string
  readonly body: readonly DocumentNode[]
}

export interface ExportListItem {
  readonly label: string
  readonly value: number
}

export interface ExportSource {
  readonly doc: DocumentNode
  /** In text order; footnotes before endnotes at the end of the export. */
  readonly notes: readonly ExportNote[]
  readonly noteOf: ReadonlyMap<DocumentNode, ExportNote>
  readonly listOf: ReadonlyMap<DocumentNode, ListInfo>
  readonly itemOf: ReadonlyMap<DocumentNode, ExportListItem>
  /** Bookmarks some internal link points to (`#name`). */
  readonly linkTargets: ReadonlySet<string>
}

const JSON_READER: ListTreeReader<DocumentNode> = {
  typeOf: (node) => node.type,
  attrsOf: (node) => node.attrs ?? {},
  childrenOf: (node) => node.content ?? [],
}

export type Mark = NonNullable<DocumentNode['marks']>[number]

const hasMark = (node: DocumentNode, type: string): boolean =>
  node.marks?.some((mark) => mark.type === type) ?? false

function blockRevisionKind(value: unknown): string | null {
  if (typeof value !== 'object' || value === null) return null
  const kind = (value as { kind?: unknown }).kind
  return kind === 'ins' || kind === 'del' ? kind : null
}

const NEEDS_BLOCK = new Set(['tableCell', 'tableHeader', 'listItem', 'noteRef', 'blockquote'])

const MERGEABLE = new Set(['paragraph', 'heading'])

/** Comment ends go to no export. */
const DROPPED = new Set(['commentStart', 'commentEnd'])

/**
 * What "Accept all" would do, without going through the editor. Comments are dropped, except in ODT
 * (`keepComments`), which carries them as annotations.
 */
export function finalDocument(node: DocumentNode, keepComments = false): DocumentNode {
  let content = node.content === undefined ? undefined : finalChildren(node.content, keepComments)
  // A block the schema does not allow empty (cell, item, note) gets a paragraph back when
  // everything in it was deleted.
  if (content?.length === 0 && NEEDS_BLOCK.has(node.type)) content = [{ type: 'paragraph' }]
  const marks = node.marks?.filter((mark) => mark.type !== 'insertion')
  const { content: _content, marks: _marks, ...rest } = node
  void _content
  void _marks
  return {
    ...rest,
    ...(content === undefined ? {} : { content }),
    ...(marks === undefined || marks.length === 0 ? {} : { marks }),
  }
}

function finalChildren(children: readonly DocumentNode[], keepComments: boolean): DocumentNode[] {
  const kept = children
    .filter((child) => survivesFinal(child, keepComments))
    .map((child) => finalDocument(child, keepComments))
  return mergeDeletedParagraphMarks(kept)
}

function survivesFinal(child: DocumentNode, keepComments: boolean): boolean {
  if (!keepComments && DROPPED.has(child.type)) return false
  if (hasMark(child, 'deletion')) return false
  return !(child.type === 'tableRow' && blockRevisionKind(child.attrs?.['rowRevision']) === 'del')
}

/** A deleted paragraph mark: its text continues in the next paragraph. */
function mergeDeletedParagraphMarks(kept: readonly DocumentNode[]): DocumentNode[] {
  const merged: DocumentNode[] = []
  let pending: DocumentNode | null = null
  for (const child of kept) {
    if (pending !== null && MERGEABLE.has(child.type)) {
      merged.push({ ...child, content: [...(pending.content ?? []), ...(child.content ?? [])] })
      pending = null
      continue
    }
    if (pending !== null) merged.push(pending)
    pending = hasDeletedParagraphMark(child) ? child : null
    if (pending === null) merged.push(child)
  }
  if (pending !== null) merged.push(pending)
  return merged
}

function hasDeletedParagraphMark(node: DocumentNode): boolean {
  return MERGEABLE.has(node.type) && blockRevisionKind(node.attrs?.['markRevision']) === 'del'
}

/** Pre-order, the text order. */
export function walk(node: DocumentNode, visit: (node: DocumentNode) => void): void {
  visit(node)
  for (const child of node.content ?? []) walk(child, visit)
}

function exportNoteOf(
  node: DocumentNode,
  section: number,
  ordinals: Record<NoteKind, number>,
  label: ReturnType<typeof noteCounter>,
): ExportNote {
  const kind = node.attrs?.['kind'] === NoteKind.Endnote ? NoteKind.Endnote : NoteKind.Footnote
  const mark = typeof node.attrs?.['mark'] === 'string' ? node.attrs['mark'] : null
  ordinals[kind] += 1
  return {
    kind,
    label: label({ kind, mark, section }),
    id: `${kind === NoteKind.Endnote ? 'nota-fim' : 'nota-rodape'}-${ordinals[kind]}`,
    body: node.content ?? [],
  }
}

function internalTargetsOf(node: DocumentNode): string[] {
  return (node.marks ?? [])
    .map((mark) => linkHref(mark))
    .filter((href): href is string => href?.startsWith('#') === true)
    .map((href) => href.slice(1))
}

export function prepareExport(
  model: Pick<DocumentModel, 'doc' | 'notes'>,
  options: { readonly keepComments?: boolean } = {},
): ExportSource {
  const doc = finalDocument(model.doc, options.keepComments === true)

  const noteOf = new Map<DocumentNode, ExportNote>()
  const notes: ExportNote[] = []
  const label = noteCounter(model.notes)
  const ordinals = { [NoteKind.Footnote]: 0, [NoteKind.Endnote]: 0 }
  let section = 0
  const listNodes: DocumentNode[] = []
  const itemNodes: DocumentNode[] = []
  const linkTargets = new Set<string>()

  walk(doc, (node) => {
    if (node.type === 'noteRef') {
      const note = exportNoteOf(node, section, ordinals, label)
      noteOf.set(node, note)
      notes.push(note)
    }
    if (typeof node.attrs?.['sectionBreak'] === 'string') section += 1
    if (LIST_TYPES.includes(node.type)) listNodes.push(node)
    if (node.type === 'listItem') itemNodes.push(node)
    for (const target of internalTargetsOf(node)) linkTargets.add(target)
  })

  // `numberLists` counts in the same pre-order.
  const numbering = numberLists(doc, JSON_READER)
  const listOf = new Map<DocumentNode, ListInfo>()
  listNodes.forEach((node, index) => {
    const info = numbering.lists[index]
    if (info !== undefined) listOf.set(node, info)
  })
  const itemOf = new Map<DocumentNode, ExportListItem>()
  itemNodes.forEach((node, index) => {
    itemOf.set(node, { label: numbering.labels[index] ?? '', value: numbering.values[index] ?? index + 1 })
  })

  const ordered = [
    ...notes.filter((note) => note.kind === NoteKind.Footnote),
    ...notes.filter((note) => note.kind === NoteKind.Endnote),
  ]
  return { doc, notes: ordered, noteOf, listOf, itemOf, linkTargets }
}

function linkHref(mark: Mark): string | null {
  if (mark.type !== 'link') return null
  const href = mark.attrs?.['href']
  return typeof href === 'string' ? href : null
}

/**
 * An internal anchor or a protocol the editor opens; `javascript:`, `file:` and `data:` become
 * text.
 */
export function safeHref(mark: Mark): string | null {
  const href = linkHref(mark)?.trim()
  if (href === undefined || href === null || href === '') return null
  if (href.startsWith('#')) return href.length > 1 ? href : null
  try {
    const parsed = new URL(href)
    return (ALLOWED_EXTERNAL_PROTOCOLS as readonly string[]).includes(parsed.protocol) ? href : null
  } catch {
    return null
  }
}

/** `null` if it is not a `data:` image. */
export function imageData(src: unknown): { readonly mime: string; readonly base64: string } | null {
  if (typeof src !== 'string') return null
  const match = /^data:(image\/[a-z0-9.+-]+);base64,([a-z0-9+/=\s]+)$/i.exec(src)
  if (match === null) return null
  return { mime: match[1]!.toLowerCase(), base64: match[2]!.replace(/\s+/g, '') }
}

const EXTENSIONS: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/jpg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
  'image/bmp': 'bmp',
  'image/svg+xml': 'svg',
  'image/tiff': 'tif',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
  'image/emf': 'emf',
  'image/wmf': 'wmf',
}

export function imageExtension(mime: string): string {
  return EXTENSIONS[mime] ?? 'bin'
}

const MAX_EXPORTED_HEADING_LEVEL = 6

/** HTML and Markdown only have six heading levels. */
export function exportHeadingLevel(node: DocumentNode): number {
  return Math.min(MAX_EXPORTED_HEADING_LEVEL, Math.max(1, Number(node.attrs?.['level']) || 1))
}

/** By style: `TOC2`, `toc 2`, `Sumário2`. */
export function tocLevelOf(node: DocumentNode): number {
  const style = node.attrs?.['styleId']
  const match = typeof style === 'string' ? /(\d)\s*$/.exec(style) : null
  return match === null ? 1 : Math.max(1, Number(match[1]))
}

/** Without the page number: on a web page it would point nowhere. */
export function withoutPageNumbers(content: readonly DocumentNode[]): DocumentNode[] {
  const kept = content.filter(
    (node) => !(node.type === 'field' && /^\s*PAGEREF\b/i.test(String(node.attrs?.['instr'] ?? ''))),
  )
  // The tab that separated the title from the page number is left over at the end.
  while (kept.length > 0) {
    const last = kept[kept.length - 1]!
    if (last.type !== 'text' || (last.text ?? '').trimEnd() !== '') break
    kept.pop()
  }
  const last = kept[kept.length - 1]
  if (last?.type === 'text' && last.text !== undefined) {
    kept[kept.length - 1] = { ...last, text: last.text.replace(/\t\s*[0-9ivxlcdm]*\s*$/i, '').trimEnd() }
  }
  return kept
}

export function exportTitle(model: Pick<DocumentModel, 'properties'>, fileName: string): string {
  const title = model.properties?.title?.trim()
  if (title !== undefined && title !== '') return title
  const dot = fileName.lastIndexOf('.')
  return dot > 0 ? fileName.slice(0, dot) : fileName
}

/** A paragraph that only carries a section mark has no text to export. */
export function isSectionMarkOnly(node: DocumentNode): boolean {
  return node.attrs?.['sectionMark'] === true && (node.content ?? []).length === 0
}
