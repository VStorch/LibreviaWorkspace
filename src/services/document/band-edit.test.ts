import { describe, expect, it } from 'vitest'
import { DEFAULT_PAGE_SETUP, type PageSetup } from './model.js'
import { editBandFloat, editBandPiece, type Band, type BandPiece } from './band.js'
import type { FloatingObject } from './floating.js'

/** Text typed in the band goes into `page`, where the writer reads it from. */
describe('editBandPiece', () => {
  const piece = (text: string, pid?: string): BandPiece => ({
    kind: 'text',
    text,
    bold: false,
    italic: false,
    ...(pid === undefined ? {} : { pid }),
  })

  const band = (...pieces: BandPiece[]): Band => ({
    left: pieces,
    center: [],
    right: [],
    rule: false,
    floats: [],
    rows: [],
  })

  const setup = (bands: Partial<PageSetup>): PageSetup => ({ ...DEFAULT_PAGE_SETUP, ...bands })

  it('troca o texto da peça daquele endereço', () => {
    const page = setup({ headerBand: band(piece('Chamado 10001', 'rId5:0:0'), piece('Título', 'rId5:1:0')) })

    expect(editBandPiece(page, 'rId5:1:0', 'Outro título').headerBand?.left.map((item) => item.text)).toEqual(
      ['Chamado 10001', 'Outro título'],
    )
  })

  it('alcança a peça esteja ela onde estiver na faixa', () => {
    // A corporate header is a table, and the text to change lives in a cell, not in one of the
    // three columns.
    const page = setup({
      firstFooterBand: {
        ...band(),
        rows: [
          { cells: [{ pieces: [piece('Mês/ANO', 'rId7:2:0')], width: 1, span: 1, rowSpan: 1, borders: '' }] },
        ],
      },
    })

    const updated = editBandPiece(page, 'rId7:2:0', 'Setembro/2026')
    expect(updated.firstFooterBand?.rows[0]?.cells[0]?.pieces[0]?.text).toBe('Setembro/2026')
  })

  it('devolve a mesma configuração quando não há o que trocar', () => {
    // A click that changed nothing must not mark the document as modified, and leaving the piece
    // without typing is the common case.
    const page = setup({ headerBand: band(piece('Chamado 10001', 'rId5:0:0')) })

    expect(editBandPiece(page, 'rId5:0:0', 'Chamado 10001')).toBe(page)
    expect(editBandPiece(page, 'rId9:4:0', 'Outra coisa')).toBe(page)
  })

  it('não escreve em peça sem endereço', () => {
    // Page numbers and images have no `w:t` to hold typed text.
    const page = setup({ headerBand: band(piece('5')) })

    expect(editBandPiece(page, 'rId5:0:0', 'Outro')).toBe(page)
  })
})

/** The whole box goes into `page`. */
describe('editBandFloat', () => {
  const caixa = (bid: string | undefined, text: string): FloatingObject => ({
    kind: 'text',
    content: [{ type: 'paragraph', content: [{ type: 'text', text }] }],
    widthMm: 80,
    heightMm: 10,
    rotation: 0,
    hFrom: 'page',
    vFrom: 'paragraph',
    behind: false,
    wrap: 'none',
    ...(bid === undefined ? {} : { bid }),
  })

  const setup = (floats: FloatingObject[]): PageSetup => ({
    ...DEFAULT_PAGE_SETUP,
    headerBand: { left: [], center: [], right: [], rule: false, floats, rows: [] },
  })

  const texto = (page: PageSetup, at: number): unknown =>
    page.headerBand?.floats[at]?.content?.[0]?.content?.[0]?.text

  it('troca o conteúdo da caixa daquele endereço', () => {
    const page = setup([caixa('rId13#0', 'EVIDÊNCIAS DO ROTEIRO'), caixa('rId13#1', 'Outra caixa')])
    const novo = [{ type: 'paragraph', content: [{ type: 'text', text: 'EVIDÊNCIAS DE HOMOLOGAÇÃO' }] }]

    const updated = editBandFloat(page, 'rId13#0', novo)
    expect(texto(updated, 0)).toBe('EVIDÊNCIAS DE HOMOLOGAÇÃO')
    expect(texto(updated, 1)).toBe('Outra caixa')
  })

  it('não escreve em caixa sem endereço', () => {
    // A box carrying numbering loses its address when its bullet changes: what is on screen is this
    // sheet's number, and writing it back would replace the `PAGE` field with a fixed number.
    const page = setup([caixa(undefined, '3')])

    expect(editBandFloat(page, 'rId13#0', [])).toBe(page)
  })

  it('devolve a mesma configuração quando o texto não mudou', () => {
    const page = setup([caixa('rId13#0', 'Igual')])
    const mesmo = [{ type: 'paragraph', content: [{ type: 'text', text: 'Igual' }] }]

    expect(editBandFloat(page, 'rId13#0', mesmo)).toBe(page)
  })
})
