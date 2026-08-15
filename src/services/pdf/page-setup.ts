import {
  PageOrientation,
  PageSize,
  hasBandContent,
  type Band,
  type BandPiece,
  type PageSetup,
} from '@services/document/model.js'

/**
 * Tradução da configuração de página para as opções do `printToPDF`.
 *
 * O Chromium trabalha em **polegadas**; a interface e o modelo trabalham em
 * milímetros. Converter no lugar errado produz um PDF com margens
 * silenciosamente erradas, então a conversão mora aqui, sozinha e testada.
 */

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

/**
 * Espaço mínimo, em milímetros, para que o cabeçalho ou rodapé caiba.
 *
 * O Chromium desenha os dois *dentro* da margem e recorta o que passar. Com
 * margem apertada, o texto simplesmente não aparece — e o usuário não teria
 * como saber por quê.
 */
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
 * Converte o texto do cabeçalho/rodapé no HTML que o Chromium espera.
 *
 * `{n}` vira o número da página e `{total}` o total — o Chromium substitui os
 * elementos com as classes `pageNumber` e `totalPages`. O texto do usuário é
 * escapado: ele não pode injetar marcação no template.
 *
 * A fonte precisa ser declarada explicitamente; sem isso o Chromium renderiza
 * o template em tamanho zero e nada aparece.
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
 * Faixa preservada do documento → HTML do template do Chromium.
 *
 * A imagem entra como `data:` URI. O template do `printToPDF` roda num
 * contexto isolado que **não busca recurso externo nenhum** — logotipo por URL
 * simplesmente não apareceria. Embutido, aparece.
 *
 * A escala é reduzida de propósito: o Chromium renderiza o template com uma
 * escala própria, e um logotipo no tamanho declarado sai maior no papel do que
 * na tela.
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
    // O tamanho declarado no documento é para a página inteira; no template do
    // Chromium ele sai desproporcional, então entra reduzido.
    (piece.fontSize === undefined ? '' : `font-size:${scaleFontSize(piece.fontSize)};`)

  return `<span style="${style}">${escapeHtml(piece.text ?? '')}</span>`
}

function scaleFontSize(fontSize: string): string {
  const value = Number.parseFloat(fontSize)
  return Number.isFinite(value) ? `${(value * 0.6).toFixed(1)}pt` : fontSize
}

/** Opções do diálogo nativo de impressão. */
export interface NativePrintOptions {
  readonly pageSize: 'A4' | 'Letter'
  readonly landscape: boolean
  readonly printBackground: boolean
  readonly margins: { readonly marginType: 'custom' } & PdfMargins
}

/**
 * Opções para `webContents.print()`.
 *
 * Cuidado que custa caro: as margens do `print()` são em **pixels CSS**, e as
 * do `printToPDF()` em **polegadas**. Usar a conversão de um no outro produz
 * margens erradas por um fator de 96 — e o erro só aparece no papel.
 */
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

/** 1 polegada = 96 pixels CSS. */
export function mmToPixels(mm: number): number {
  return (mm / MM_PER_INCH) * 96
}

export function buildPrintOptions(page: PageSetup): PdfPrintOptions {
  // A faixa preservada do documento manda: quando ela existe, é o cabeçalho
  // real do arquivo, com logotipo e numeração. O texto digitado só vale para
  // documentos criados aqui.
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
