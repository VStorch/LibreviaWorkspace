import { describe, expect, it } from 'vitest'
import { DEFAULT_PAGE_SETUP, editBandPiece, type Band, type BandPiece, type PageSetup } from './model.js'

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

    expect(editBandPiece(page, 'rId5:1:0', 'Outro título').headerBand?.left.map((item) => item.text)).toEqual([
      'Chamado 10001',
      'Outro título',
    ])
  })

  it('alcança a peça esteja ela onde estiver na faixa', () => {
    // Cabeçalho corporativo é uma tabela, e o texto que se quer trocar mora
    // numa célula — não numa das três colunas.
    const page = setup({
      firstFooterBand: {
        ...band(),
        rows: [{ cells: [{ pieces: [piece('Mês/ANO', 'rId7:2:0')], width: 1, span: 1, rowSpan: 1, borders: '' }] }],
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
