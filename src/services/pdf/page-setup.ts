import { PageOrientation, PageSize, type PageSetup } from '@services/document/model.js'

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
  const hasHeader = page.header.trim().length > 0
  const hasFooter = page.footer.trim().length > 0

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
    headerTemplate: buildHeaderFooterTemplate(page.header),
    footerTemplate: buildHeaderFooterTemplate(page.footer),
    // As margens e o tamanho vêm daqui, não do CSS: é uma fonte só.
    preferCSSPageSize: false,
    scale: 1,
  }
}
