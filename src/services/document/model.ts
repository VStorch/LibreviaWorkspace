/**
 * `doc` is the ProseMirror JSON Tiptap edits; `page` and the rest carry what does not fit in the
 * text flow.
 */

import { NO_BANDS, type Band, type BandHeights } from './band.js'
import { BUILTIN_STYLES, type StyleSheet } from './styles.js'

export const PageSize = {
  A4: 'A4',
  Letter: 'Letter',
} as const
export type PageSize = (typeof PageSize)[keyof typeof PageSize]

export const PageOrientation = {
  Portrait: 'portrait',
  Landscape: 'landscape',
} as const
export type PageOrientation = (typeof PageOrientation)[keyof typeof PageOrientation]

/** In millimetres, the unit the UI shows. */
export interface Margins {
  readonly top: number
  readonly right: number
  readonly bottom: number
  readonly left: number
}

export interface PageSetup {
  readonly size: PageSize
  readonly orientation: PageOrientation
  readonly margins: Margins
  /** Plain text typed by the user; `{n}` and `{total}` are replaced when rendering the PDF. */
  readonly header: string
  readonly footer: string
  /** The imported document's header; when present, it wins on screen. */
  readonly headerBand: Band | null
  readonly footerBand: Band | null
  /**
   * Only exist when the document turns on `w:titlePg` or `w:evenAndOddHeaders`: Word keeps the
   * parts even with them off; see `PageReader.HasTitlePage`.
   */
  readonly firstHeaderBand: Band | null
  readonly firstFooterBand: Band | null
  readonly evenHeaderBand: Band | null
  readonly evenFooterBand: Band | null
  /** `w:pgMar/@header`: the vertical origin for anchors inside the header. */
  readonly headerDistanceMm: number
  readonly footerDistanceMm: number
  /** `w:pgNumType`. Absent means "decimal from 1" on screen and "leave alone" when saving. */
  readonly pageNumberFormat?: PageNumberFormat | undefined
  readonly pageNumberStart?: number | null | undefined
  /** `w:titlePg` and `w:evenAndOddHeaders`. Absent, they follow what the bands say. */
  readonly titlePage?: boolean | null | undefined
  readonly evenAndOddHeaders?: boolean | null | undefined
  /** `w:sectPr/w:type`. Absent means "next page" on screen and "leave alone" when saving. */
  readonly start?: SectionStart | undefined
  /** `w:cols`. Absent means one column on screen and "leave alone" when saving. */
  readonly columns?: SectionColumns | undefined
}

/** `w:cols`: column count, spacing between them (mm) and the separator line. */
export interface SectionColumns {
  readonly count: number
  readonly spaceMm: number
  readonly separator: boolean
  /**
   * Unequal widths, as the file declares them (`w:equalWidth="0"`). The screen draws equal columns;
   * changing columns in the panel equalizes them.
   */
  readonly widthsMm?: number[] | undefined
}

/** OOXML section starts, from `w:type/@w:val`. */
export const SECTION_STARTS = ['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'] as const
export type SectionStart = (typeof SECTION_STARTS)[number]

/**
 * The paragraph that **closes** the section carries the same `id` in `sectionBreak`, as OOXML keeps
 * `w:sectPr` in the paragraph that closes the section. The last section is `DocumentModel.page`. A
 * null band from the second section on means "link to previous"; see `effectiveSections`.
 */
export interface SectionSetup extends PageSetup {
  readonly id: string
}

/** The ones the panel offers, from `w:pgNumType/@w:fmt`. */
export const PAGE_NUMBER_FORMATS = [
  'decimal',
  'lowerRoman',
  'upperRoman',
  'lowerLetter',
  'upperLetter',
] as const
export type PageNumberFormat = (typeof PAGE_NUMBER_FORMATS)[number]

/**
 * Mutable collections because Tiptap expects `JSONContent`, and a `readonly` array is not
 * assignable to a plain one.
 */
export interface DocumentNode {
  readonly type: string
  readonly content?: DocumentNode[]
  readonly text?: string
  readonly attrs?: Record<string, unknown>
  readonly marks?: { readonly type: string; readonly attrs?: Record<string, unknown> }[]
}

export interface DocumentModel {
  /** The last section: the body's, and the only one in a single-section document. */
  readonly page: PageSetup
  /** The ones before the last, in order. Outside the nodes for the same reason as `styles`. */
  readonly sections?: readonly SectionSetup[]
  readonly doc: DocumentNode
  /**
   * Outside the nodes: a block fingerprint is made from what is inside it, and a style there would
   * make every block look changed to the surgical save.
   */
  readonly styles: StyleSheet
  /**
   * `.sdoc` draft < 4, with the effective formatting on each block. Saving compares it with a
   * flattened reading of the original, otherwise every block would look changed.
   */
  readonly flattened?: boolean
  /**
   * `.sdoc` draft < 5: no bookmark, field, internal link or table of contents in the nodes. Same
   * reason as `flattened`.
   */
  readonly beforeReferences?: boolean
  /** `.sdoc` draft < 6: no `sectionBreak`. Same reason as `flattened`. */
  readonly beforeSections?: boolean
  /**
   * File bookmarks that did not become nodes: between table rows, loose between blocks, in the
   * header or in a box. A reference citing them is not broken, and "Update fields" leaves its
   * result as Word left it.
   */
  readonly outsideBookmarks?: readonly string[]
  /**
   * The body lives outside the nodes, like `styles`; the text only holds the anchor ends, one per
   * thread. Only those the text supports (`resolveComments`).
   */
  readonly comments?: readonly DocumentComment[]
  /** `.sdoc` draft < 7: no anchor in the nodes. Same reason as `flattened`. */
  readonly beforeComments?: boolean
  /** `w:trackRevisions`. Absent means "leave alone". */
  readonly trackChanges?: boolean
  /** `.sdoc` draft < 8: no revision marks. Same reason as `flattened`. */
  readonly beforeRevisions?: boolean
  /**
   * Outside the nodes, like `styles`: a note's number is the order of its reference. Absent means
   * Word's numbering: 1, 2, 3 for footnotes and i, ii, iii for endnotes.
   */
  readonly notes?: DocumentNotes
  /** `.sdoc` draft < 9: no `noteRef`. Same reason as `flattened`. */
  readonly beforeNotes?: boolean
  /** `.sdoc` draft < 11: no `math`. Same reason as `flattened`. */
  readonly beforeMath?: boolean
  /**
   * `docProps/core.xml` and part of `app.xml`, outside the nodes. When saving each field is a
   * patch: absent means "leave as is", empty means "clear".
   */
  readonly properties?: DocumentProperties
}

/**
 * As Word shows them in File → Properties. Dates are W3CDTF (`2026-10-02T12:00:00Z`), as in the
 * package.
 */
export interface DocumentProperties {
  readonly title?: string
  readonly subject?: string
  /** `dc:creator`: the author(s), separated by semicolons, as in Word. */
  readonly creator?: string
  readonly keywords?: string
  readonly category?: string
  /** `dc:description`: what Word calls Comments. */
  readonly description?: string
  readonly lastModifiedBy?: string
  readonly revision?: string
  readonly created?: string
  readonly modified?: string
  /** `docProps/app.xml`. */
  readonly company?: string
  readonly manager?: string
  /** `TotalTime` from `app.xml`, in minutes. Read only: the editor does not measure it. */
  readonly totalTime?: number
}

/** `w:footnotePr`/`w:endnotePr`. */
export interface NoteNumbering {
  /** `decimal`, `lowerRoman`, `upperLetter`, `chicago`… */
  readonly numFmt?: string
  readonly start?: number
  /** `continuous`, `eachSect`, `eachPage`. */
  readonly restart?: string
  readonly pos?: string
}

export interface DocumentNotes {
  readonly footnotePr?: NoteNumbering
  readonly endnotePr?: NoteNumbering
}

/** As `word/comments.xml` and `word/commentsExtended.xml` describe it. */
export interface DocumentComment {
  /** The `w:id`, the same `cid` as the anchor ends in the text. */
  readonly id: string
  /** The comment this one replies to. Absent for the one opening the thread. */
  readonly parentId?: string
  readonly author: string
  readonly initials?: string
  /** As the file has it (ISO 8601); empty when it has none. */
  readonly date: string
  /** Each paragraph's text, unformatted. */
  readonly paragraphs: readonly string[]
  /** `w15:done`; applies to the thread, through the comment that opens it. */
  readonly done: boolean
  readonly paraId?: string
  /** The body has formatting, an image or a field plain text does not show. */
  readonly rich?: boolean
}

export const PAGE_DIMENSIONS_MM: Record<PageSize, { width: number; height: number }> = {
  [PageSize.A4]: { width: 210, height: 297 },
  [PageSize.Letter]: { width: 216, height: 279 },
}

/** Like Word's "Normal" default: 2.54 cm all around. */
export const DEFAULT_PAGE_SETUP: PageSetup = {
  size: PageSize.A4,
  orientation: PageOrientation.Portrait,
  margins: { top: 25, right: 25, bottom: 25, left: 25 },
  header: '',
  footer: '',
  headerBand: null,
  footerBand: null,
  firstHeaderBand: null,
  firstFooterBand: null,
  evenHeaderBand: null,
  evenFooterBand: null,
  headerDistanceMm: 12.5,
  footerDistanceMm: 12.5,
}

export const EMPTY_DOCUMENT: DocumentNode = {
  type: 'doc',
  content: [{ type: 'paragraph' }],
}

export function createEmptyDocument(): DocumentModel {
  return { page: DEFAULT_PAGE_SETUP, doc: EMPTY_DOCUMENT, styles: BUILTIN_STYLES }
}

/** Already accounting for orientation. */
export function pageDimensionsMm(page: PageSetup): { width: number; height: number } {
  const base = PAGE_DIMENSIONS_MM[page.size]
  return page.orientation === PageOrientation.Landscape
    ? { width: base.height, height: base.width }
    : { width: base.width, height: base.height }
}

/** The width that sets the frame on screen. */
export function contentWidthMm(page: PageSetup): number {
  const { width } = pageDimensionsMm(page)
  return width - page.margins.left - page.margins.right
}

export function contentHeightMm(page: PageSetup, bands: BandHeights = NO_BANDS): number {
  const { height } = pageDimensionsMm(page)
  const inset = contentInsetsMm(page, bands)
  return height - inset.top - inset.bottom
}

/**
 * The margin is a floor: when the header is taller than its distance to the edge plus the margin,
 * Word and LibreOffice push the body below it.
 */
export function contentInsetsMm(page: PageSetup, bands: BandHeights): { top: number; bottom: number } {
  return {
    top: Math.max(page.margins.top, page.headerDistanceMm + bands.headerMm),
    bottom: Math.max(page.margins.bottom, page.footerDistanceMm + bands.footerMm),
  }
}

/** Margins adding up to more than the page would give a negative text area. */
export function isValidMargins(page: PageSetup): boolean {
  const { width, height } = pageDimensionsMm(page)
  const values = [page.margins.top, page.margins.right, page.margins.bottom, page.margins.left]

  if (values.some((value) => !Number.isFinite(value) || value < 0)) return false
  return page.margins.left + page.margins.right < width && page.margins.top + page.margins.bottom < height
}
