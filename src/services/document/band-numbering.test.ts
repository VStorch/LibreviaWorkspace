import { describe, expect, it } from 'vitest'
import { DEFAULT_PAGE_SETUP, type PageSetup } from './model.js'
import { appendPageField, bandForPage, pageLabel, pieceText, plainBand, type Band } from './band.js'

const band = (text: string, pid?: string): Band => ({
  left: [],
  center: [{ kind: 'text', text, bold: false, italic: false, ...(pid === undefined ? {} : { pid }) }],
  right: [],
  rule: false,
  floats: [],
  rows: [],
})

const textOf = (found: Band | null): string | undefined => found?.center[0]?.text

describe('page numbering', () => {
  it('the printed number follows the start and format of w:pgNumType', () => {
    const page: PageSetup = { ...DEFAULT_PAGE_SETUP, pageNumberFormat: 'lowerRoman', pageNumberStart: 3 }
    expect([1, 2, 3].map((sheet) => pageLabel(page, sheet))).toEqual(['iii', 'iv', 'v'])
    expect(pageLabel(DEFAULT_PAGE_SETUP, 2)).toBe('2')
  })

  it("the text piece shows {n} and {total} with this sheet's numbers", () => {
    const piece = { kind: 'text' as const, text: 'Folha {n} de {total}', bold: false, italic: false }
    expect(pieceText(piece, 'iv', 9)).toBe('Folha iv de 9')
    // Braces written in the file are text, not a field.
    expect(pieceText({ ...piece, literal: true }, 'iv', 9)).toBe('Folha {n} de {total}')
  })

  it('different first page without its own band leaves the cover clean', () => {
    const page: PageSetup = { ...DEFAULT_PAGE_SETUP, footerBand: band('Rodapé'), titlePage: true }
    expect(bandForPage(page, 1, 'footer')).toBeNull()
    expect(textOf(bandForPage(page, 2, 'footer'))).toBe('Rodapé')
  })

  it('the cover band saved with the switch off does not show', () => {
    const page: PageSetup = {
      ...DEFAULT_PAGE_SETUP,
      footerBand: band('Miolo'),
      firstFooterBand: band('Capa'),
      titlePage: false,
    }
    expect(textOf(bandForPage(page, 1, 'footer'))).toBe('Miolo')
  })

  it('an old draft without the switch counts by the band it brought', () => {
    const page: PageSetup = {
      ...DEFAULT_PAGE_SETUP,
      footerBand: band('Miolo'),
      firstFooterBand: band('Capa'),
    }
    expect(textOf(bandForPage(page, 1, 'footer'))).toBe('Capa')
  })

  it('parity is that of the printed number, not of the sheet', () => {
    const page: PageSetup = {
      ...DEFAULT_PAGE_SETUP,
      footerBand: band('Ímpar'),
      evenFooterBand: band('Par'),
      evenAndOddHeaders: true,
      pageNumberStart: 2,
    }
    expect(textOf(bandForPage(page, 1, 'footer'))).toBe('Par')
    expect(textOf(bandForPage(page, 2, 'footer'))).toBe('Ímpar')
  })

  it('the plain text footer becomes a band, with the field in place of {n}', () => {
    const footer = plainBand('Página {n} de {total}')!
    expect(footer.center.map((piece) => piece.kind)).toEqual(['text', 'pageNumber', 'text', 'totalPages'])
    const page: PageSetup = { ...DEFAULT_PAGE_SETUP, footer: 'Página {n}' }
    expect(bandForPage(page, 1, 'footer')?.center[1]?.kind).toBe('pageNumber')
    expect(plainBand('   ')).toBeNull()
  })

  it('inserting the field with no cursor in the band goes to the end of the footer', () => {
    expect(appendPageField(DEFAULT_PAGE_SETUP, '{n}')?.footer).toBe('{n}')
    expect(appendPageField({ ...DEFAULT_PAGE_SETUP, footer: 'Página' }, '{n}')?.footer).toBe('Página {n}')

    const fromWord: PageSetup = { ...DEFAULT_PAGE_SETUP, footerBand: band('Manual', 'p1') }
    expect(appendPageField(fromWord, '{total}')?.footerBand?.center[0]?.text).toBe('Manual {total}')

    const noText: PageSetup = { ...DEFAULT_PAGE_SETUP, footerBand: band('Logotipo') }
    expect(appendPageField(noText, '{n}')).toBeNull()
  })
})
