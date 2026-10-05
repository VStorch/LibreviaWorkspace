import { describe, expect, it } from 'vitest'
import { linesOf, type BandPiece } from './band.js'

/**
 * Each band paragraph is a line, like the three lines of the manual template footer, stacked and
 * centered in LibreOffice.
 */
describe('linesOf', () => {
  const piece = (text: string, line = false): BandPiece => ({
    kind: 'text',
    text,
    bold: false,
    italic: false,
    line,
  })

  it('quebra onde o arquivo abre parágrafo', () => {
    const lines = linesOf([
      piece('www.exemplo.com.br'),
      piece('Documento V01', true),
      piece(' - Fulano'),
      piece('Mês/ANO', true),
    ])

    expect(lines.map((line) => line.map((item) => item.text))).toEqual([
      ['www.exemplo.com.br'],
      ['Documento V01', ' - Fulano'],
      ['Mês/ANO'],
    ])
  })

  it('sem marca nenhuma, tudo é uma linha só', () => {
    expect(linesOf([piece('Relatório'), piece(' mensal')])).toHaveLength(1)
  })

  it('faixa vazia não gera linha', () => {
    expect(linesOf([])).toEqual([])
  })
})
