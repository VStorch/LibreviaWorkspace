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

describe('numeração de página', () => {
  it('o número impresso segue o início e o formato de w:pgNumType', () => {
    const page: PageSetup = { ...DEFAULT_PAGE_SETUP, pageNumberFormat: 'lowerRoman', pageNumberStart: 3 }
    expect([1, 2, 3].map((sheet) => pageLabel(page, sheet))).toEqual(['iii', 'iv', 'v'])
    expect(pageLabel(DEFAULT_PAGE_SETUP, 2)).toBe('2')
  })

  it('a peça de texto mostra {n} e {total} com os números desta folha', () => {
    const piece = { kind: 'text' as const, text: 'Folha {n} de {total}', bold: false, italic: false }
    expect(pieceText(piece, 'iv', 9)).toBe('Folha iv de 9')
  })

  it('primeira página diferente sem faixa própria deixa a capa limpa', () => {
    const page: PageSetup = { ...DEFAULT_PAGE_SETUP, footerBand: band('Rodapé'), titlePage: true }
    expect(bandForPage(page, 1, 'footer')).toBeNull()
    expect(textOf(bandForPage(page, 2, 'footer'))).toBe('Rodapé')
  })

  it('a faixa da capa guardada com o interruptor desligado não aparece', () => {
    const page: PageSetup = {
      ...DEFAULT_PAGE_SETUP,
      footerBand: band('Miolo'),
      firstFooterBand: band('Capa'),
      titlePage: false,
    }
    expect(textOf(bandForPage(page, 1, 'footer'))).toBe('Miolo')
  })

  it('rascunho antigo, sem o interruptor, vale pela faixa que trouxe', () => {
    const page: PageSetup = {
      ...DEFAULT_PAGE_SETUP,
      footerBand: band('Miolo'),
      firstFooterBand: band('Capa'),
    }
    expect(textOf(bandForPage(page, 1, 'footer'))).toBe('Capa')
  })

  it('a paridade é a do número impresso, e não a da folha', () => {
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

  it('o rodapé de texto simples vira faixa, com o campo no lugar de {n}', () => {
    const footer = plainBand('Página {n} de {total}')!
    expect(footer.center.map((piece) => piece.kind)).toEqual(['text', 'pageNumber', 'text', 'totalPages'])
    const page: PageSetup = { ...DEFAULT_PAGE_SETUP, footer: 'Página {n}' }
    expect(bandForPage(page, 1, 'footer')?.center[1]?.kind).toBe('pageNumber')
    expect(plainBand('   ')).toBeNull()
  })

  it('inserir o campo sem cursor na faixa vai para o fim do rodapé', () => {
    expect(appendPageField(DEFAULT_PAGE_SETUP, '{n}')?.footer).toBe('{n}')
    expect(appendPageField({ ...DEFAULT_PAGE_SETUP, footer: 'Página' }, '{n}')?.footer).toBe('Página {n}')

    const fromWord: PageSetup = { ...DEFAULT_PAGE_SETUP, footerBand: band('Manual', 'p1') }
    expect(appendPageField(fromWord, '{total}')?.footerBand?.center[0]?.text).toBe('Manual {total}')

    const noText: PageSetup = { ...DEFAULT_PAGE_SETUP, footerBand: band('Logotipo') }
    expect(appendPageField(noText, '{n}')).toBeNull()
  })
})
