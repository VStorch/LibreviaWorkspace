import { describe, expect, it } from 'vitest'
import { findOccurrences, stepIndex } from './search.js'

describe('findOccurrences', () => {
  it('finds all occurrences', () => {
    expect(findOccurrences('boi boi boi', 'boi')).toEqual([
      { start: 0, end: 3 },
      { start: 4, end: 7 },
      { start: 8, end: 11 },
    ])
  })

  it('ignores case by default', () => {
    expect(findOccurrences('Contrato contrato CONTRATO', 'contrato')).toHaveLength(3)
  })

  it('respects case when requested', () => {
    expect(findOccurrences('Contrato contrato', 'contrato', true)).toEqual([{ start: 9, end: 17 }])
  })

  it('does not return overlapping occurrences', () => {
    // Otherwise "replace all" would reprocess the newly written text and might never end.
    expect(findOccurrences('aaaa', 'aa')).toEqual([
      { start: 0, end: 2 },
      { start: 2, end: 4 },
    ])
  })

  it('returns empty for an empty term', () => {
    expect(findOccurrences('qualquer texto', '')).toEqual([])
  })

  it('returns empty when there is no occurrence', () => {
    expect(findOccurrences('texto', 'ausente')).toEqual([])
  })

  it('keeps positions aligned when a letter changes length in lowercase', () => {
    // `'İ'.toLowerCase()` has two code units: lowercasing the whole text would shift the following
    // positions.
    const text = 'İstanbul tem contrato'
    const [occurrence] = findOccurrences(text, 'contrato')

    expect(occurrence).toBeDefined()
    expect(text.slice(occurrence!.start, occurrence!.end)).toBe('contrato')
  })

  it('keeps positions aligned after a non-BMP character', () => {
    const text = '😀 Contrato'
    const [occurrence] = findOccurrences(text, 'contrato')

    expect(text.slice(occurrence!.start, occurrence!.end)).toBe('Contrato')
  })

  it('treats accents as ordinary characters', () => {
    // "acao" does not find "ação": normalizing accents would change positions and break
    // replacement. If ever wanted, it needs an index map, not a naive normalization.
    expect(findOccurrences('ação', 'acao')).toEqual([])
    expect(findOccurrences('ação e ação', 'ação')).toHaveLength(2)
  })
})

describe('stepIndex', () => {
  it('advances and wraps to the start past the end', () => {
    expect(stepIndex(0, 3, 1)).toBe(1)
    expect(stepIndex(2, 3, 1)).toBe(0)
  })

  it('steps back and wraps to the end past the start', () => {
    expect(stepIndex(0, 3, -1)).toBe(2)
  })

  it('starts from the first occurrence when there is no selection yet', () => {
    expect(stepIndex(-1, 3, 1)).toBe(0)
    expect(stepIndex(-1, 3, -1)).toBe(2)
  })

  it('returns -1 when there are no occurrences', () => {
    expect(stepIndex(0, 0, 1)).toBe(-1)
  })
})
