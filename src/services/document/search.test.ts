import { describe, expect, it } from 'vitest'
import { findOccurrences, stepIndex } from './search.js'

describe('findOccurrences', () => {
  it('encontra todas as ocorrências', () => {
    expect(findOccurrences('boi boi boi', 'boi')).toEqual([
      { start: 0, end: 3 },
      { start: 4, end: 7 },
      { start: 8, end: 11 },
    ])
  })

  it('ignora maiúsculas por padrão', () => {
    expect(findOccurrences('Contrato contrato CONTRATO', 'contrato')).toHaveLength(3)
  })

  it('respeita maiúsculas quando pedido', () => {
    expect(findOccurrences('Contrato contrato', 'contrato', true)).toEqual([{ start: 9, end: 17 }])
  })

  it('não devolve ocorrências sobrepostas', () => {
    // Otherwise "replace all" would reprocess the newly written text and might never end.
    expect(findOccurrences('aaaa', 'aa')).toEqual([
      { start: 0, end: 2 },
      { start: 2, end: 4 },
    ])
  })

  it('devolve vazio para termo vazio', () => {
    expect(findOccurrences('qualquer texto', '')).toEqual([])
  })

  it('devolve vazio quando não há ocorrência', () => {
    expect(findOccurrences('texto', 'ausente')).toEqual([])
  })

  it('mantém as posições alinhadas quando a letra muda de tamanho em minúsculas', () => {
    // `'İ'.toLowerCase()` has two code units: lowercasing the whole text would shift the following
    // positions.
    const texto = 'İstanbul tem contrato'
    const [ocorrencia] = findOccurrences(texto, 'contrato')

    expect(ocorrencia).toBeDefined()
    expect(texto.slice(ocorrencia!.start, ocorrencia!.end)).toBe('contrato')
  })

  it('mantém as posições alinhadas depois de um caractere fora do BMP', () => {
    const texto = '😀 Contrato'
    const [ocorrencia] = findOccurrences(texto, 'contrato')

    expect(texto.slice(ocorrencia!.start, ocorrencia!.end)).toBe('Contrato')
  })

  it('trata acentuação como caractere comum', () => {
    // "acao" does not find "ação": normalizing accents would change positions and break
    // replacement. If ever wanted, it needs an index map, not a naive normalization.
    expect(findOccurrences('ação', 'acao')).toEqual([])
    expect(findOccurrences('ação e ação', 'ação')).toHaveLength(2)
  })
})

describe('stepIndex', () => {
  it('avança e volta ao início ao passar do fim', () => {
    expect(stepIndex(0, 3, 1)).toBe(1)
    expect(stepIndex(2, 3, 1)).toBe(0)
  })

  it('retrocede e vai ao fim ao passar do início', () => {
    expect(stepIndex(0, 3, -1)).toBe(2)
  })

  it('parte da primeira ocorrência quando ainda não há seleção', () => {
    expect(stepIndex(-1, 3, 1)).toBe(0)
    expect(stepIndex(-1, 3, -1)).toBe(2)
  })

  it('devolve -1 quando não há ocorrências', () => {
    expect(stepIndex(0, 0, 1)).toBe(-1)
  })
})
