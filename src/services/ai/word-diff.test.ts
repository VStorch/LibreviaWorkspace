import { describe, expect, it } from 'vitest'
import { DiffOp, wordDiff } from './word-diff.js'

/** Removed in brackets, added in braces. */
const shown = (before: string, after: string): string =>
  wordDiff(before, after)
    .map(({ op, text }) => (op === DiffOp.Removed ? `[${text}]` : op === DiffOp.Added ? `{${text}}` : text))
    .join('')

describe('wordDiff', () => {
  it('marks only the words that changed', () => {
    expect(shown('o time fico contente em marco', 'o time ficou contente em março')).toBe(
      'o time [fico]{ficou} contente em [marco]{março}',
    )
  })

  it('punctuation and spaces count apart from the words', () => {
    expect(shown('Olá mundo', 'Olá, mundo!')).toBe('Olá{,} mundo{!}')
  })

  it('a removal comes before the addition at the same place', () => {
    expect(shown('um dois três', 'um quatro três')).toBe('um [dois]{quatro} três')
  })

  it('all new, all gone, or nothing changed', () => {
    expect(shown('', 'novo')).toBe('{novo}')
    expect(shown('velho', '')).toBe('[velho]')
    expect(wordDiff('igual', 'igual')).toEqual([{ op: DiffOp.Same, text: 'igual' }])
  })

  it('a text too long to compare word by word changes whole', () => {
    const long = 'palavra '.repeat(600)
    expect(wordDiff(long, `${long}fim`)).toEqual([
      { op: DiffOp.Removed, text: long },
      { op: DiffOp.Added, text: `${long}fim` },
    ])
  })
})
