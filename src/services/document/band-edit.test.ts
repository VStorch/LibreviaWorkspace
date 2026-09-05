import { describe, expect, it } from 'vitest'
import { DEFAULT_PAGE_SETUP, type PageSetup } from './model.js'
import { editBandFloat, editBandPiece, type Band, type BandPiece } from './band.js'
import type { FloatingObject } from './floating.js'

/**
 * O texto digitado na faixa entra na configuração de página.
 *
 * A faixa não mora no documento do editor: ela é a parte OOXML preservada, e
 * vive em `page`. Quem a edita mexe aqui, e é daqui que o gravador a lê.
 */
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
    // Cabeçalho corporativo é uma tabela, e o texto que se quer trocar mora
    // numa célula — não numa das três colunas.
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
    // Um clique que não mudou nada não pode marcar o documento como alterado, e
    // sair da peça sem digitar é o caso comum.
    const page = setup({ headerBand: band(piece('Chamado 10001', 'rId5:0:0')) })

    expect(editBandPiece(page, 'rId5:0:0', 'Chamado 10001')).toBe(page)
    expect(editBandPiece(page, 'rId9:4:0', 'Outra coisa')).toBe(page)
  })

  it('não escreve em peça sem endereço', () => {
    // Número de página e imagem não têm `w:t` onde guardar o que se digitasse.
    const page = setup({ headerBand: band(piece('5')) })

    expect(editBandPiece(page, 'rId5:0:0', 'Outro')).toBe(page)
  })
})

/**
 * O texto de uma caixa da faixa entra na configuração de página.
 *
 * O cabeçalho corporativo não é feito de parágrafos soltos: é um grupo de
 * formas, e o título mora dentro de uma caixa. Ela vem inteira porque digitar
 * dentro dela abre e fecha parágrafos.
 */
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
    // A caixa que traz numeração perde o endereço ao ter o marcador trocado: o
    // que está na tela é o número desta folha, e devolvê-lo trocaria o campo
    // `PAGE` por um número fixo.
    const page = setup([caixa(undefined, '3')])

    expect(editBandFloat(page, 'rId13#0', [])).toBe(page)
  })

  it('devolve a mesma configuração quando o texto não mudou', () => {
    const page = setup([caixa('rId13#0', 'Igual')])
    const mesmo = [{ type: 'paragraph', content: [{ type: 'text', text: 'Igual' }] }]

    expect(editBandFloat(page, 'rId13#0', mesmo)).toBe(page)
  })
})
