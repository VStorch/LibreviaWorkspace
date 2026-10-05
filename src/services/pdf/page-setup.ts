import { PageOrientation, PageSize, type PageSetup } from '@services/document/model.js'
import { hasBandContent, type Band, type BandPiece } from '@services/document/band.js'
import { mmToInches, mmToPx } from '@services/units.js'

export interface PdfMargins {
  readonly top: number
  readonly bottom: number
  readonly left: number
  readonly right: number
}

export interface PdfPrintOptions {
  readonly pageSize: 'A4' | 'Letter'
  readonly landscape: boolean
  readonly margins: PdfMargins
  readonly printBackground: boolean
  readonly displayHeaderFooter: boolean
  readonly headerTemplate: string
  readonly footerTemplate: string
  readonly preferCSSPageSize: boolean
  readonly scale: number
}

/** Chromium draws header and footer inside the margin and clips what overflows. */
export const MIN_MARGIN_FOR_HEADER_MM = 12

export function marginFitsHeaderOrFooter(marginMm: number): boolean {
  return marginMm >= MIN_MARGIN_FOR_HEADER_MM
}

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/**
 * Chromium replaces the `pageNumber` and `totalPages` classes. The text is escaped, and the font is
 * declared because without it the template renders at size zero.
 */
export function buildHeaderFooterTemplate(text: string): string {
  if (text.trim().length === 0) return '<span></span>'

  const withTokens = escapeHtml(text)
    .replaceAll('{n}', '<span class="pageNumber"></span>')
    .replaceAll('{total}', '<span class="totalPages"></span>')

  return (
    '<div style="font-family: Carlito, Calibri, sans-serif; font-size: 9pt; color: #444; ' +
    'width: 100%; padding: 0 12mm; box-sizing: border-box; text-align: center;">' +
    withTokens +
    '</div>'
  )
}

/**
 * The image goes in as a `data:` URI because the template fetches no external resource, and the
 * scale is reduced because Chromium draws it at its own scale.
 */
export function buildBandTemplate(band: Band): string {
  const cell = (pieces: readonly BandPiece[], align: string): string =>
    `<div style="display:flex;align-items:center;gap:6px;justify-content:${align};white-space:nowrap;min-width:0">` +
    pieces.map(pieceToHtml).join('') +
    '</div>'

  const rule = band.rule ? 'border-bottom:1px solid #999;padding-bottom:2px;' : ''

  return (
    '<div style="font-family: Carlito, Calibri, sans-serif; font-size: 8pt; color: #222; ' +
    `width: 100%; padding: 0 8mm; box-sizing: border-box; display: grid; ` +
    `grid-template-columns: auto 1fr auto; align-items: center; gap: 8px; ${rule}">` +
    cell(band.left, 'flex-start') +
    cell(band.center, 'center') +
    cell(band.right, 'flex-end') +
    '</div>'
  )
}

function pieceToHtml(piece: BandPiece): string {
  if (piece.kind === 'image') {
    if (piece.src === undefined) return ''
    const width = piece.width === undefined ? '' : `width:${Math.round(piece.width * 0.75)}px;`
    return `<img src="${escapeHtml(piece.src)}" style="${width}object-fit:contain" />`
  }

  if (piece.kind === 'pageNumber') return '<span class="pageNumber"></span>'
  if (piece.kind === 'totalPages') return '<span class="totalPages"></span>'

  const style =
    (piece.bold ? 'font-weight:700;' : '') +
    (piece.italic ? 'font-style:italic;' : '') +
    (piece.color === undefined ? '' : `color:${escapeHtml(piece.color)};`) +
    (piece.fontSize === undefined ? '' : `font-size:${scaleFontSize(piece.fontSize)};`)

  return `<span style="${style}">${escapeHtml(piece.text ?? '')}</span>`
}

function scaleFontSize(fontSize: string): string {
  const value = Number.parseFloat(fontSize)
  return Number.isFinite(value) ? `${(value * 0.6).toFixed(1)}pt` : fontSize
}

export interface NativePrintOptions {
  readonly pageSize: 'A4' | 'Letter'
  readonly landscape: boolean
  readonly printBackground: boolean
  readonly margins: { readonly marginType: 'custom' } & PdfMargins
}

/** `print()` margins are CSS pixels, `printToPDF()` ones are inches. */
export function buildNativePrintOptions(page: PageSetup): NativePrintOptions {
  return {
    pageSize: page.size === PageSize.Letter ? 'Letter' : 'A4',
    landscape: page.orientation === PageOrientation.Landscape,
    printBackground: true,
    margins: {
      marginType: 'custom',
      top: Math.round(mmToPx(page.margins.top)),
      bottom: Math.round(mmToPx(page.margins.bottom)),
      left: Math.round(mmToPx(page.margins.left)),
      right: Math.round(mmToPx(page.margins.right)),
    },
  }
}

/**
 * @param paged The HTML already comes in paper-sized sheets: zero margin, `@page` size and no
 * Chromium band, which would fall over ours. A spreadsheet is not paginated.
 */
export function buildPrintOptions(page: PageSetup, paged = false): PdfPrintOptions {
  if (paged) {
    return {
      pageSize: page.size === PageSize.Letter ? 'Letter' : 'A4',
      landscape: page.orientation === PageOrientation.Landscape,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      printBackground: true,
      displayHeaderFooter: false,
      headerTemplate: '',
      footerTemplate: '',
      // The same paper the screen drew.
      preferCSSPageSize: true,
      scale: 1,
    }
  }

  // The document's preserved band wins; typed text only applies without one.
  const headerBand = hasBandContent(page.headerBand) ? page.headerBand : null
  const footerBand = hasBandContent(page.footerBand) ? page.footerBand : null

  const hasHeader = headerBand !== null || page.header.trim().length > 0
  const hasFooter = footerBand !== null || page.footer.trim().length > 0

  return {
    pageSize: page.size === PageSize.Letter ? 'Letter' : 'A4',
    landscape: page.orientation === PageOrientation.Landscape,
    margins: {
      top: mmToInches(page.margins.top),
      bottom: mmToInches(page.margins.bottom),
      left: mmToInches(page.margins.left),
      right: mmToInches(page.margins.right),
    },
    // Otherwise text highlight and table header backgrounds vanish from the PDF.
    printBackground: true,
    displayHeaderFooter: hasHeader || hasFooter,
    headerTemplate:
      headerBand === null ? buildHeaderFooterTemplate(page.header) : buildBandTemplate(headerBand),
    footerTemplate:
      footerBand === null ? buildHeaderFooterTemplate(page.footer) : buildBandTemplate(footerBand),
    // Margins and size come from here, not from CSS: a single source.
    preferCSSPageSize: false,
    scale: 1,
  }
}
