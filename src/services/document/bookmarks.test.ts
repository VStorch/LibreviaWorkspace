import { describe, expect, it } from 'vitest'
import { hiddenBookmarkName, isHiddenBookmark, isValidBookmarkName, nextBookmarkId } from './bookmarks.js'

describe('isValidBookmarkName', () => {
  it('aceita letra seguida de letras, algarismos e sublinhado', () => {
    expect(isValidBookmarkName('Resumo')).toBe(true)
    expect(isValidBookmarkName('Seção_2')).toBe(true)
  })

  it('recusa espaço, algarismo no começo, sublinhado no começo e nome longo', () => {
    expect(isValidBookmarkName('Meu resumo')).toBe(false)
    expect(isValidBookmarkName('2a')).toBe(false)
    expect(isValidBookmarkName('_Toc1')).toBe(false)
    expect(isValidBookmarkName('a'.repeat(41))).toBe(false)
    expect(isValidBookmarkName('a'.repeat(40))).toBe(true)
  })
})

describe('marcadores ocultos e ids', () => {
  it('oculto é o que começa por sublinhado', () => {
    expect(isHiddenBookmark('_Toc100')).toBe(true)
    expect(isHiddenBookmark('Resumo')).toBe(false)
  })

  it('o id novo é um a mais que o maior número, e ignora o que não é número', () => {
    expect(nextBookmarkId([])).toBe('0')
    expect(nextBookmarkId(['0', '7', 'x', '3'])).toBe('8')
  })

  it('o nome oculto novo conta a partir do maior do mesmo prefixo', () => {
    expect(hiddenBookmarkName('_Ref', ['_Ref000000004', '_Toc900', 'Resumo'])).toBe('_Ref000000005')
    expect(hiddenBookmarkName('_Toc', [])).toBe('_Toc000000001')
  })
})
