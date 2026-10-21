import { describe, expect, it } from 'vitest'
import { contextWindow, stripSelectionTags } from './context.js'

describe('contextWindow', () => {
  it('sends everything when it fits the budget', () => {
    expect(contextWindow('antes ', ' depois', 100)).toEqual({
      before: 'antes ',
      after: ' depois',
      excerpt: false,
    })
  })

  it('splits the budget in half and cuts at a word boundary', () => {
    const window = contextWindow('um dois três quatro', 'cinco seis sete oito', 16)
    expect(window).toEqual({ before: 'quatro', after: 'cinco', excerpt: true })
  })

  it('what one side does not use goes to the other', () => {
    const window = contextWindow('', 'um dois três quatro cinco', 14)
    expect(window.after).toBe('um dois três')
    expect(window.excerpt).toBe(true)
  })

  it('without a budget, sends no context', () => {
    expect(contextWindow('antes', 'depois', 0)).toEqual({ before: '', after: '', excerpt: false })
  })

  it('strips selection marks written in the document itself', () => {
    expect(stripSelectionTags('a <selection>b</SELECTION> c')).toBe('a b c')
    expect(contextWindow('x <selection>', '', 100).before).toBe('x ')
  })

  it('does not let a selection mark form after cleaning nor pass with spaces', () => {
    expect(stripSelectionTags('a<sel<selection>ection>b')).toBe('ab')
    expect(stripSelectionTags('<sel<sel<selection>ection>ection>')).toBe('')
    expect(stripSelectionTags('a<selection >b</ selection>c< SELECTION x>d')).toBe('abcd')
  })
})
