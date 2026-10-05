import type { DocumentNode, PageSetup } from './model.js'
import { formatNumber } from './list-numbering.js'
import type { FloatingObject } from './floating.js'

/**
 * A preserved header or footer. Text of pieces with an address is editable and goes back to the
 * `w:t` it came from; the rest of the OOXML part goes back untouched.
 */
export interface Band {
  readonly left: BandPiece[]
  readonly center: BandPiece[]
  readonly right: BandPiece[]
  readonly rule: boolean
  /** What does not fit in three columns: a drawing with a real position, possibly rotated. */
  readonly floats: FloatingObject[]
  /** The grid, when the header is a table: logo in a merged cell, title beside it. */
  readonly rows: BandRow[]
}

export interface BandRow {
  readonly cells: BandCell[]
}

export interface BandCell {
  readonly pieces: BandPiece[]
  /** Fraction of the grid width, 0 to 1. */
  readonly width: number
  readonly span: number
  readonly rowSpan: number
  readonly align?: string | undefined
  /** Initials of the sides with a line: `t`, `l`, `b`, `r`. */
  readonly borders: string
}

export interface BandPiece {
  readonly kind: 'text' | 'image' | 'pageNumber' | 'totalPages'
  // Explicit `| undefined` because of `exactOptionalPropertyTypes`: this type must be assignable to
  // what zod infers in the shared schema.
  readonly text?: string | undefined
  readonly src?: string | undefined
  readonly width?: number | undefined
  readonly height?: number | undefined
  readonly bold: boolean
  readonly italic: boolean
  readonly color?: string | undefined
  readonly fontSize?: string | undefined
  /** A CSS font stack, as the reader resolved it. */
  readonly fontFamily?: string | undefined
  /** Headers and footers are paragraphs, and each paragraph is a line. */
  readonly line?: boolean | undefined
  /**
   * Where the piece lives in the file: saving writes into its `w:t` and leaves the rest of the
   * header alone. Page numbers, images and tabs have no address.
   */
  readonly pid?: string | undefined
  /**
   * The file text already had `{n}` or `{total}` written. It is text: the screen does not replace
   * it with the number, and saving does not turn it into a field.
   */
  readonly literal?: boolean | undefined
}

/** Shared by screen and paper, so the two drawings do not drift. */
export function linesOf(pieces: readonly BandPiece[]): BandPiece[][] {
  const lines: BandPiece[][] = []
  for (const piece of pieces) {
    if (piece.line === true || lines.length === 0) lines.push([])
    lines[lines.length - 1]!.push(piece)
  }

  return lines
}

/** Measured on the sheet: only exists after drawing. */
export interface BandHeights {
  readonly headerMm: number
  readonly footerMm: number
}

export const NO_BANDS: BandHeights = { headerMm: 0, footerMm: 0 }

/**
 * Word's order: the title page beats parity, and parity beats the default. A missing band falls
 * back to the default; `hasBandContent` decides whether the sheet stays blank.
 */
export function bandForPage(page: PageSetup, sheet: number, kind: 'header' | 'footer'): Band | null {
  const first = kind === 'header' ? page.firstHeaderBand : page.firstFooterBand
  const even = kind === 'header' ? page.evenHeaderBand : page.evenFooterBand
  const fallback =
    (kind === 'header' ? page.headerBand : page.footerBand) ??
    plainBand(kind === 'header' ? page.header : page.footer)

  // Switched on without its own band means a blank sheet, as in Word: that is how "Different first
  // page" gives a cover without a number.
  if (sheet === 1 && usesTitlePage(page)) return first
  // Parity follows the printed number, not the sheet: starting at 2, the first sheet is already
  // even.
  if (usesEvenAndOdd(page) && pageNumberOf(page, sheet) % 2 === 0) return even
  return fallback
}

/** "Different first page": what the document says, or what the bands imply. */
export function usesTitlePage(page: PageSetup): boolean {
  return page.titlePage ?? (page.firstHeaderBand !== null || page.firstFooterBand !== null)
}

/** "Different odd and even pages", by the same criterion. */
export function usesEvenAndOdd(page: PageSetup): boolean {
  return page.evenAndOddHeaders ?? (page.evenHeaderBand !== null || page.evenFooterBand !== null)
}

/** The number printed on sheet `sheet` (from 1): the `w:pgNumType` start plus the offset. */
export function pageNumberOf(page: PageSetup, sheet: number): number {
  return (page.pageNumberStart ?? 1) + sheet - 1
}

/** As the `PAGE` field writes it, in the `w:pgNumType` format. */
export function pageLabel(page: PageSetup, sheet: number): string {
  return formatNumber(pageNumberOf(page, sheet), page.pageNumberFormat ?? 'decimal')
}

/**
 * The text may carry `{n}` and `{total}` until saving turns them into fields
 * (`BandWriter.Rewrite`).
 */
export function pieceText(piece: BandPiece, label: string, total: number): string {
  if (piece.kind === 'pageNumber') return label
  if (piece.kind === 'totalPages') return String(total)
  if (piece.literal === true) return piece.text ?? ''
  return substituteFields(piece.text ?? '', label, total)
}

export function substituteFields(text: string, label: string, total: number): string {
  return text.replaceAll('{n}', label).replaceAll('{total}', String(total))
}

/** `TemplateStyles.BandFont`, the font the file receives. */
const PLAIN_BAND_FONT = 'Calibri, Carlito, sans-serif'

/**
 * The "Page setup" line as a band, so the paginated sheet draws it as the file stores it: centered,
 * 9 pt grey.
 */
export function plainBand(text: string): Band | null {
  if (text.trim().length === 0) return null
  const style = { bold: false, italic: false, color: '#444444', fontSize: '9pt', fontFamily: PLAIN_BAND_FONT }
  const center: BandPiece[] = []
  for (const part of text.split(/(\{n\}|\{total\})/)) {
    if (part === '') continue
    if (part === '{n}') center.push({ kind: 'pageNumber', ...style })
    else if (part === '{total}') center.push({ kind: 'totalPages', ...style })
    else center.push({ kind: 'text', text: part, ...style })
  }
  return { left: [], center, right: [], rule: false, floats: [], rows: [] }
}

/**
 * The file has a single header: changing the piece updates every sheet, as in Word. Returns the
 * same setup when nothing changes, so the document is not marked dirty.
 */
export function editBandPiece<T extends PageSetup>(page: T, pid: string, text: string): T {
  let changed = false

  const inPieces = (pieces: BandPiece[]): BandPiece[] =>
    pieces.map((piece) => {
      if (piece.pid !== pid || piece.text === text) return piece
      changed = true
      return { ...piece, text }
    })

  const updated = mapBands(page, (band) => ({
    ...band,
    left: inPieces(band.left),
    center: inPieces(band.center),
    right: inPieces(band.right),
    rows: band.rows.map((row) => ({
      cells: row.cells.map((cell) => ({ ...cell, pieces: inPieces(cell.pieces) })),
    })),
  }))

  return changed ? updated : page
}

/** The whole box comes in: typing inside it opens and closes paragraphs. */
export function editBandFloat<T extends PageSetup>(page: T, bid: string, content: DocumentNode[]): T {
  let changed = false

  const updated = mapBands(page, (band) => ({
    ...band,
    floats: band.floats.map((object) => {
      if (object.bid !== bid) return object
      if (JSON.stringify(object.content ?? []) === JSON.stringify(content)) return object
      changed = true
      return { ...object, content }
    }),
  }))

  return changed ? updated : page
}

export function hasBandContent(band: Band | null): band is Band {
  return (
    band !== null &&
    (band.left.length > 0 ||
      band.center.length > 0 ||
      band.right.length > 0 ||
      band.rows.length > 0 ||
      band.rule)
  )
}

/** The six bands: title page, even and default, for header and footer. */
function mapBands<T extends PageSetup>(page: T, transform: (band: Band) => Band): T {
  const at = (band: Band | null): Band | null => (band === null ? null : transform(band))

  return {
    ...page,
    headerBand: at(page.headerBand),
    footerBand: at(page.footerBand),
    firstHeaderBand: at(page.firstHeaderBand),
    firstFooterBand: at(page.firstFooterBand),
    evenHeaderBand: at(page.evenHeaderBand),
    evenFooterBand: at(page.evenFooterBand),
  }
}

/** Half the margin: the band is wider than the text column, as in corporate documents. */
export function bandInsetMm(page: PageSetup): number {
  return Math.min(page.margins.left, page.margins.right) / 2
}

/**
 * Without the cursor in a band, the field goes to the end of the footer, as in Word. In a file
 * footer without editable text it returns `null`: inventing a paragraph in Word's part is what the
 * surgical save does not do.
 */
export function appendPageField(page: PageSetup, token: '{n}' | '{total}'): PageSetup | null {
  const band = page.footerBand
  if (band === null || !hasBandContent(band)) {
    const footer = page.footer.trimEnd()
    return { ...page, footer: footer === '' ? token : `${footer} ${token}` }
  }

  const candidates = [
    ...band.rows.flatMap((row) => row.cells.flatMap((cell) => cell.pieces)),
    ...band.left,
    ...band.center,
    ...band.right,
  ].filter((piece) => piece.kind === 'text' && piece.pid !== undefined)
  const target = candidates.at(-1)
  if (target === undefined || target.pid === undefined) return null

  const text = target.text ?? ''
  const joined = text === '' || text.endsWith(' ') ? `${text}${token}` : `${text} ${token}`
  return editBandPiece(page, target.pid, joined)
}
