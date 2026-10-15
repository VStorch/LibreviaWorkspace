import { describe, expect, it } from 'vitest'
import { contextWindow, stripSelectionTags } from './context.js'

describe('contextWindow', () => {
  it('manda tudo quando cabe no orçamento', () => {
    expect(contextWindow('antes ', ' depois', 100)).toEqual({
      before: 'antes ',
      after: ' depois',
      excerpt: false,
    })
  })

  it('divide o orçamento ao meio e corta em limite de palavra', () => {
    const window = contextWindow('um dois três quatro', 'cinco seis sete oito', 16)
    expect(window).toEqual({ before: 'quatro', after: 'cinco', excerpt: true })
  })

  it('o que um lado não usa vai para o outro', () => {
    const window = contextWindow('', 'um dois três quatro cinco', 14)
    expect(window.after).toBe('um dois três')
    expect(window.excerpt).toBe(true)
  })

  it('sem orçamento não manda contexto', () => {
    expect(contextWindow('antes', 'depois', 0)).toEqual({ before: '', after: '', excerpt: false })
  })

  it('tira as marcas de seleção escritas no próprio documento', () => {
    expect(stripSelectionTags('a <selection>b</SELECTION> c')).toBe('a b c')
    expect(contextWindow('x <selection>', '', 100).before).toBe('x ')
  })

  it('não deixa marca de seleção se formar depois de limpar nem passar com espaços', () => {
    expect(stripSelectionTags('a<sel<selection>ection>b')).toBe('ab')
    expect(stripSelectionTags('<sel<sel<selection>ection>ection>')).toBe('')
    expect(stripSelectionTags('a<selection >b</ selection>c< SELECTION x>d')).toBe('abcd')
  })
})
