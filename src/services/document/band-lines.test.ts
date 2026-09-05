import { describe, expect, it } from 'vitest'
import { linesOf, type BandPiece } from './band.js'

/**
 * Cada parágrafo do cabeçalho ou do rodapé é uma linha.
 *
 * O rodapé do modelo de manual tem três — endereço, autoria e data. Emendados,
 * viravam uma frase que atravessava a folha, enquanto o LibreOffice mostrava as
 * três empilhadas e centradas.
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
