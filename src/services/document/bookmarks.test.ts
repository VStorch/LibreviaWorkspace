import { describe, expect, it } from 'vitest'
import { hiddenBookmarkName, isHiddenBookmark, isValidBookmarkName, nextBookmarkId } from './bookmarks.js'

describe('isValidBookmarkName', () => {
  it('accepts a letter followed by letters, digits and underscores', () => {
    expect(isValidBookmarkName('Resumo')).toBe(true)
    expect(isValidBookmarkName('Seção_2')).toBe(true)
  })

  it('rejects space, leading digit, leading underscore and long name', () => {
    expect(isValidBookmarkName('Meu resumo')).toBe(false)
    expect(isValidBookmarkName('2a')).toBe(false)
    expect(isValidBookmarkName('_Toc1')).toBe(false)
    expect(isValidBookmarkName('a'.repeat(41))).toBe(false)
    expect(isValidBookmarkName('a'.repeat(40))).toBe(true)
  })
})

describe('hidden bookmarks and ids', () => {
  it('hidden means starting with an underscore', () => {
    expect(isHiddenBookmark('_Toc100')).toBe(true)
    expect(isHiddenBookmark('Resumo')).toBe(false)
  })

  it('the new id is one more than the largest number, and ignores non-numbers', () => {
    expect(nextBookmarkId([])).toBe('0')
    expect(nextBookmarkId(['0', '7', 'x', '3'])).toBe('8')
  })

  it('the new hidden name counts from the largest with the same prefix', () => {
    expect(hiddenBookmarkName('_Ref', ['_Ref000000004', '_Toc900', 'Resumo'])).toBe('_Ref000000005')
    expect(hiddenBookmarkName('_Toc', [])).toBe('_Toc000000001')
  })
})
