import { PageOrientation, PageSize, type PageSetup } from '@services/document/model.js'
import { hasBandContent, type Band, type BandPiece } from '@services/document/band.js'

/** O Chromium trabalha em polegadas, e o modelo em milímetros: a conversão mora aqui. */

export const MM_PER_INCH = 25.4

export function mmToInches(mm: number): number {
  return mm / MM_PER_INCH
}

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

/** O Chromium desenha cabeçalho e rodapé dentro da margem e recorta o que passar. */
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
 * O Chromium troca as classes `pageNumber` e `totalPages`. O texto é escapado, e
 * a fonte é declarada porque sem ela o template sai em tamanho zero.
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
 * A imagem entra como `data:` URI porque o template não busca recurso externo,
 * e a escala é reduzida porque o Chromium o desenha com escala própria.
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

/** As margens do `print()` são em pixels CSS, e as do `printToPDF()` em polegadas. */
export function buildNativePrintOptions(page: PageSetup): NativePrintOptions {
  return {
    pageSize: page.size === PageSize.Letter ? 'Letter' : 'A4',
    landscape: page.orientation === PageOrientation.Landscape,
    printBackground: true,
    margins: {
      marginType: 'custom',
      top: Math.round(mmToPixels(page.margins.top)),
      bottom: Math.round(mmToPixels(page.margins.bottom)),
      left: Math.round(mmToPixels(page.margins.left)),
      right: Math.round(mmToPixels(page.margins.right)),
    },
  }
}

export function mmToPixels(mm: number): number {
  return (mm / MM_PER_INCH) * 96
}

/**
 * @param paged
 * O HTML já vem em folhas do tamanho do papel: margem zero, tamanho do `@page`
 * e sem a faixa do Chromium, que cairia por cima da nossa. A planilha não vem
 * paginada.
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
      // O mesmo papel que a tela desenhou.
      preferCSSPageSize: true,
      scale: 1,
    }
  }

  // A faixa preservada do documento manda; o texto digitado só vale sem ela.
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
    // Sem isto, destaque de texto e fundo de cabeçalho de tabela somem do PDF.
    printBackground: true,
    displayHeaderFooter: hasHeader || hasFooter,
    headerTemplate:
      headerBand === null ? buildHeaderFooterTemplate(page.header) : buildBandTemplate(headerBand),
    footerTemplate:
      footerBand === null ? buildHeaderFooterTemplate(page.footer) : buildBandTemplate(footerBand),
    // As margens e o tamanho vêm daqui, não do CSS: é uma fonte só.
    preferCSSPageSize: false,
    scale: 1,
  }
}
