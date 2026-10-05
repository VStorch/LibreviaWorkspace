import type { Band, BandPiece } from './band.js'
import type { SheetPlan } from './paginate.js'
import {
  contentWidthMm,
  type DocumentNode,
  type PageSetup,
  type SectionSetup,
  type SectionStart,
} from './model.js'
import { INDENT_STEP_MM } from '@services/units.js'

/** Half an inch, the column spacing Word suggests. */
export const DEFAULT_COLUMN_SPACING_MM = INDENT_STEP_MM

/**
 * Each section stores only the bands it **declares**, like the file. One that does not declare a
 * kind uses the previous one's ("Link to previous" in Word), kind by kind: title, even and default
 * inherit separately.
 */

export const BAND_KEYS = [
  'headerBand',
  'footerBand',
  'firstHeaderBand',
  'firstFooterBand',
  'evenHeaderBand',
  'evenFooterBand',
] as const
export type BandKey = (typeof BAND_KEYS)[number]

/** The earlier ones and, last, the body's. */
export function allSections(page: PageSetup, sections: readonly SectionSetup[] = []): PageSetup[] {
  return [...sections, page]
}

/**
 * From the second section on, a null band inherits; in the first it means "no band". A declared but
 * empty band does not inherit: it is the blank sheet the user wanted.
 */
export function effectiveSections(page: PageSetup, sections: readonly SectionSetup[] = []): PageSetup[] {
  const resolved: PageSetup[] = []
  for (const section of allSections(page, sections)) {
    const previous = resolved.at(-1)
    if (previous === undefined) {
      resolved.push(section)
      continue
    }
    const inherited: Partial<Record<BandKey, Band | null>> = {}
    for (const key of BAND_KEYS) {
      if (section[key] === null || section[key] === undefined) inherited[key] = previous[key]
    }
    resolved.push({ ...section, ...inherited })
  }
  return resolved
}

export function sectionBreakOf(
  node: { readonly attrs?: Record<string, unknown> | null } | DocumentNode,
): string | null {
  const value = node.attrs?.['sectionBreak']
  return typeof value === 'string' && value.length > 0 ? value : null
}

export interface SectionBlock {
  readonly attrs: Record<string, unknown>
  readonly isTextblock?: boolean
  descendants?: (callback: (node: SectionBlock) => boolean | void) => void
}

/** In a list item the mark is on the inner paragraph, and the section ends with the whole list. */
export function sectionBreakIn(block: SectionBlock): string | null {
  const own = sectionBreakOf(block)
  if (own !== null || block.isTextblock === true || block.descendants === undefined) return own
  let found: string | null = null
  block.descendants((node) => {
    if (found !== null) return false
    found = sectionBreakOf(node)
    return found === null && node.isTextblock !== true
  })
  return found
}

/**
 * Indexes into `allSections`. A block belongs to the section of the next mark, because the mark
 * **closes** the section, like `w:sectPr`. A mark with an unknown id (a paragraph pasted from
 * another document) does not count.
 */
export function blockSections(
  breaks: readonly (string | null)[],
  sections: readonly SectionSetup[],
): number[] {
  const indexOf = new Map(sections.map((section, index) => [section.id, index]))
  const result = new Array<number>(breaks.length)
  let current = sections.length
  for (let index = breaks.length - 1; index >= 0; index--) {
    const id = breaks[index]
    const known = id === null || id === undefined ? undefined : indexOf.get(id)
    if (known !== undefined) current = known
    result[index] = current
  }
  return result
}

/** Same paper and orientation: the sheet does not change size between the two. */
export function samePaper(left: PageSetup, right: PageSetup): boolean {
  return left.size === right.size && left.orientation === right.orientation
}

/** "Next page", even and odd always; continuous only when the paper changes, as in Word. */
export function startsNewSheet(section: PageSetup, previous: PageSetup | undefined): boolean {
  if (previous === undefined) return false
  const start = section.start ?? 'nextPage'
  if (start === 'continuous' || start === 'nextColumn') return !samePaper(section, previous)
  return true
}

export function parityOf(section: PageSetup): 'even' | 'odd' | null {
  return section.start === 'evenPage' ? 'even' : section.start === 'oddPage' ? 'odd' : null
}

export interface SheetSetup {
  /**
   * With `pageNumberStart` set to the section's first sheet number, so `bandForPage` and
   * `pageLabel` get it right.
   */
  readonly page: PageSetup
  /** Counting from the section's first sheet, from 1. */
  readonly inSection: number
}

/**
 * The blank sheet of an even or odd section already belongs to the new section, but is not its
 * title page.
 */
export function sheetSetups(
  effective: readonly PageSetup[],
  sheets: readonly Pick<SheetPlan, 'section' | 'number' | 'first' | 'blank'>[],
): SheetSetup[] {
  const result: SheetSetup[] = []
  let inSection = 0
  let firstNumber = 1
  let previous: number | null = null
  for (const sheet of sheets) {
    if (sheet.first || sheet.section !== previous) {
      inSection = 0
      firstNumber = sheet.number
    }
    previous = sheet.section
    inSection += 1
    const section = effective[sheet.section] ?? effective.at(-1)!
    result.push({
      page: { ...section, pageNumberStart: firstNumber, ...(sheet.blank ? { titlePage: false } : {}) },
      inSection,
    })
  }
  return result
}

export interface SectionList {
  readonly page: PageSetup
  readonly sections: readonly SectionSetup[]
}

/** `n1`, `n2`… (those read from the file are `s1`, `s2`…). */
export function freshSectionId(sections: readonly SectionSetup[]): string {
  const taken = new Set(sections.map((section) => section.id))
  let counter = sections.length + 1
  while (taken.has(`n${counter}`)) counter += 1
  return `n${counter}`
}

/**
 * As in Word: the upper section is a copy of the split one, and the lower one stays the same
 * section, starting as the break asked. The `w:type` belongs to the starting section.
 */
export function withSectionBreak(
  list: SectionList,
  index: number,
  id: string,
  start: SectionStart,
): SectionList {
  const all = allSections(list.page, list.sections)
  const current = all[index] ?? list.page
  // The copy gets the new id over the split section's (the body's has none).
  const upper: SectionSetup = { ...current, id }
  if (index >= list.sections.length) {
    return { page: { ...list.page, start }, sections: [...list.sections, upper] }
  }
  return {
    page: list.page,
    sections: [
      ...list.sections.slice(0, index),
      upper,
      { ...list.sections[index]!, start },
      ...list.sections.slice(index + 1),
    ],
  }
}

/** The break was deleted, and the range passes to the section below. */
export function withoutSection(list: SectionList, id: string): SectionList {
  return { page: list.page, sections: list.sections.filter((section) => section.id !== id) }
}

/**
 * For the whole document, the other sections get paper, margins, band distances, number format and
 * distinct title page; the numbering restart stays only where it was asked. "Odd and even" is
 * document-wide in Word.
 */
export function withPageSetup(
  list: SectionList,
  index: number,
  draft: PageSetup,
  scope: 'section' | 'document',
): SectionList {
  const shared = (section: PageSetup): Partial<PageSetup> =>
    scope === 'document'
      ? {
          size: draft.size,
          orientation: draft.orientation,
          margins: draft.margins,
          headerDistanceMm: draft.headerDistanceMm,
          footerDistanceMm: draft.footerDistanceMm,
          pageNumberFormat: draft.pageNumberFormat,
          titlePage: draft.titlePage,
          evenAndOddHeaders: draft.evenAndOddHeaders,
        }
      : { evenAndOddHeaders: draft.evenAndOddHeaders ?? section.evenAndOddHeaders }
  const last = list.sections.length
  const page = index >= last ? draft : { ...list.page, ...shared(list.page) }
  const sections = list.sections.map((section, at) =>
    at === index ? { ...draft, id: section.id } : { ...section, ...shared(section) },
  )
  return { page, sections }
}

const KIND_KEYS = {
  header: ['headerBand', 'firstHeaderBand', 'evenHeaderBand'],
  footer: ['footerBand', 'firstFooterBand', 'evenFooterBand'],
} as const satisfies Record<'header' | 'footer', readonly BandKey[]>

/** The section declares none of the three kinds on that side. */
export function isLinkedToPrevious(section: PageSetup, index: number, kind: 'header' | 'footer'): boolean {
  return index > 0 && KIND_KEYS[kind].every((key) => section[key] === null || section[key] === undefined)
}

/**
 * When unlinked, the section gets a copy of the bands it inherited, with addresses tagged by the
 * owning section (`s2~rId5:0:1`; the body's is `body`): saving creates a new part for the copy
 * (`SectionWriter.ApplyBands`).
 */
export function withBandsLinked<T extends PageSetup>(
  section: T,
  inherited: PageSetup | undefined,
  key: string,
  kind: 'header' | 'footer',
  linked: boolean,
): T {
  const changes: Partial<Record<BandKey, Band | null>> = {}
  for (const band of KIND_KEYS[kind]) {
    changes[band] = linked ? null : unlinkedCopy(inherited?.[band] ?? null, key)
  }
  return { ...section, ...changes }
}

function unlinkedCopy(band: Band | null, key: string): Band | null {
  if (band === null) return null
  const mark = (address: string | undefined): string | undefined =>
    address === undefined ? undefined : `${key}~${address.slice(address.indexOf('~') + 1)}`
  const pieces = (list: readonly BandPiece[]): BandPiece[] =>
    list.map((piece) => (piece.pid === undefined ? piece : { ...piece, pid: mark(piece.pid)! }))
  return {
    ...band,
    left: pieces(band.left),
    center: pieces(band.center),
    right: pieces(band.right),
    rows: band.rows.map((row) => ({
      cells: row.cells.map((cell) => ({ ...cell, pieces: pieces(cell.pieces) })),
    })),
    floats: band.floats.map((object) =>
      object.bid === undefined ? object : { ...object, bid: mark(object.bid)! },
    ),
  }
}

export interface ColumnGeometry {
  readonly count: number
  readonly widthMm: number
  /** From one column's left edge to the next: width plus spacing. */
  readonly stepMm: number
  readonly spaceMm: number
  readonly separator: boolean
}

/** Unequal widths (`widthsMm`) are drawn equal; the file keeps them and the inventory warns. */
export function columnGeometry(section: PageSetup): ColumnGeometry {
  const count = Math.max(1, Math.round(section.columns?.count ?? 1))
  const spaceMm = count > 1 ? Math.max(0, section.columns?.spaceMm ?? DEFAULT_COLUMN_SPACING_MM) : 0
  const content = contentWidthMm(section)
  const widthMm = Math.max((content - spaceMm * (count - 1)) / count, 1)
  return {
    count,
    widthMm,
    stepMm: widthMm + spaceMm,
    spaceMm,
    separator: section.columns?.separator === true,
  }
}

/** The panel only knows equal columns, and saving equalizes them (`SectionWriter.ApplyColumns`). */
export function withColumns(
  list: SectionList,
  index: number,
  columns: { readonly count: number; readonly spaceMm: number; readonly separator: boolean },
  scope: 'section' | 'document',
): SectionList {
  const applies = (at: number): boolean => scope === 'document' || at === index
  const last = list.sections.length
  return {
    page: applies(last) ? { ...list.page, columns: { ...columns } } : list.page,
    sections: list.sections.map((section, at) =>
      applies(at) ? { ...section, columns: { ...columns } } : section,
    ),
  }
}

/**
 * The store keeps a **library** of sections by id, and the text says which ones count: the marks in
 * body order and the `bodySection` attribute for the last. That way undo, which only knows the
 * text, also undoes the section.
 */
export interface ResolvedSections extends SectionList {
  readonly sections: readonly SectionSetup[]
  /** The library entry standing in for the last section, if any. */
  readonly bodyId: string | null
}

function withoutId(section: PageSetup): PageSetup {
  if (!('id' in section)) return section
  const { id: _id, ...rest } = section as SectionSetup
  void _id
  return rest
}

export function resolveSections(
  marks: readonly (string | null)[],
  bodyId: unknown,
  page: PageSetup,
  library: readonly SectionSetup[],
): ResolvedSections {
  const byId = new Map(library.map((section) => [section.id, section]))
  const seen = new Set<string>()
  const sections: SectionSetup[] = []
  for (const id of marks) {
    const section = id === null ? undefined : byId.get(id)
    if (section === undefined || seen.has(section.id)) continue
    seen.add(section.id)
    sections.push(section)
  }
  const body = typeof bodyId === 'string' ? byId.get(bodyId) : undefined
  return {
    page: body === undefined ? page : withoutId(body),
    sections,
    bodyId: body === undefined ? null : body.id,
  }
}

/**
 * By id and without deleting anything: an entry the text no longer uses may come back with undo.
 */
export function storeSections(
  next: SectionList,
  bodyId: string | null,
  page: PageSetup,
  library: readonly SectionSetup[],
): { page: PageSetup; library: SectionSetup[] } {
  const changed = new Map(next.sections.map((section) => [section.id, section]))
  if (bodyId !== null) changed.set(bodyId, { ...next.page, id: bodyId })
  const known = new Set(library.map((section) => section.id))
  return {
    page: bodyId === null ? next.page : page,
    library: [
      ...library.map((section) => changed.get(section.id) ?? section),
      ...[...changed.values()].filter((section) => !known.has(section.id)),
    ],
  }
}

/** Top-level block by block. */
export function marksOfJson(doc: DocumentNode): (string | null)[] {
  const find = (node: DocumentNode): string | null => {
    const own = sectionBreakOf(node)
    if (own !== null) return own
    for (const child of node.content ?? []) {
      const found = find(child)
      if (found !== null) return found
    }
    return null
  }
  return (doc.content ?? []).map(find)
}

/**
 * Only **adds** to the library, so undo works. The upper section is a copy of the split one with a
 * new id; the lower one gets another entry with the new start, pointed to by the mark closing it
 * (`rename`) or by `bodyId`.
 */
export function planSectionBreak(
  resolved: ResolvedSections,
  library: readonly SectionSetup[],
  index: number,
  start: SectionStart,
): {
  readonly upperId: string
  readonly additions: SectionSetup[]
  readonly rename: { readonly from: string; readonly to: string } | null
  readonly bodyId: string | null
} {
  const all = allSections(resolved.page, resolved.sections)
  const current = withoutId(all[index] ?? resolved.page)
  const upperId = freshSectionId(library)
  const lowerId = freshSectionId([...library, { ...current, id: upperId }])
  const lower = resolved.sections[index]
  return {
    upperId,
    additions: [
      { ...current, id: upperId },
      { ...current, id: lowerId, start },
    ],
    rename: lower === undefined ? null : { from: lower.id, to: lowerId },
    bodyId: lower === undefined ? lowerId : null,
  }
}

/**
 * Or, in the last section, the break opening it. The upper range passes to the lower section, as in
 * Word, and it receives the bands it inherited from the deleted one.
 */
export function planSectionDelete(
  resolved: ResolvedSections,
  index: number,
): { readonly removeId: string; readonly next: SectionList } | null {
  if (resolved.sections.length === 0) return null
  const target = Math.min(index, resolved.sections.length - 1)
  const removed = resolved.sections[target]!
  const effective = effectiveSections(resolved.page, resolved.sections)
  const lowerIndex = target + 1
  const all = allSections(resolved.page, resolved.sections)
  const lower = all[lowerIndex]!
  const inherited: Partial<Record<BandKey, Band | null>> = {}
  for (const key of BAND_KEYS) {
    if (lower[key] === null || lower[key] === undefined) inherited[key] = effective[target]![key] ?? null
  }
  const updated = { ...lower, ...inherited }
  const sections = resolved.sections.filter((section) => section.id !== removed.id)
  return {
    removeId: removed.id,
    next:
      lowerIndex >= resolved.sections.length
        ? { page: updated, sections }
        : {
            page: resolved.page,
            sections: sections.map((section) =>
              section.id === (lower as SectionSetup).id ? (updated as SectionSetup) : section,
            ),
          },
  }
}
