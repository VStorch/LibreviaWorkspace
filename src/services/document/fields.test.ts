import { describe, expect, it } from 'vitest'
import {
  fieldArgument,
  fieldKind,
  fieldSwitch,
  formatFieldNumber,
  sequenceNumbers,
  tocLevels,
  tocLinks,
  tocOmitsPages,
} from './fields.js'

describe('a instrução do campo', () => {
  it('dá a palavra, o argumento e as chaves, com as aspas desfeitas', () => {
    expect(fieldKind(' PAGEREF _Toc100 \\h ')).toBe('PAGEREF')
    expect(fieldKind('seq Figura')).toBe('SEQ')
    expect(fieldArgument(' REF _Ref200 \\h ')).toBe('_Ref200')
    expect(fieldArgument(' TOC \\o "1-3" ')).toBeNull()
    expect(fieldSwitch(' TOC \\o "1-3" \\h \\z ', 'o')).toBe('1-3')
    expect(fieldSwitch(' TOC \\o "1-3" \\h \\z ', 'h')).toBe('')
    expect(fieldSwitch(' TOC \\o "1-3" ', 'n')).toBeNull()
  })
})

describe('o sumário', () => {
  it('lê os níveis, o link e a omissão dos números', () => {
    expect(tocLevels(' TOC \\o "1-3" \\h \\z \\u ')).toEqual({ from: 1, to: 3 })
    expect(tocLevels(' TOC \\o "2-4" ')).toEqual({ from: 2, to: 4 })
    expect(tocLevels(' TOC \\o ')).toEqual({ from: 1, to: 9 })
    expect(tocLevels(' TOC \\t "Capítulo,1" ')).toEqual({ from: 1, to: 3 })
    expect(tocLinks(' TOC \\o "1-3" \\h ')).toBe(true)
    expect(tocLinks(' TOC \\o "1-3" ')).toBe(false)
    expect(tocOmitsPages(' TOC \\o "1-3" \\n ')).toBe(true)
  })
})

describe('a sequência', () => {
  it('conta cada identificador à parte, recomeça com \\r e repete com \\c', () => {
    expect(
      sequenceNumbers([
        ' SEQ Figura \\* ARABIC ',
        ' SEQ Tabela ',
        ' SEQ figura ',
        ' SEQ Figura \\c ',
        ' SEQ Figura \\r 7 ',
        ' SEQ Tabela \\* ROMAN ',
      ]),
    ).toEqual(['1', '1', '2', '2', '7', 'II'])
  })

  it('formata em romano e em letras', () => {
    expect(formatFieldNumber(14, 'SEQ x \\* roman')).toBe('xiv')
    expect(formatFieldNumber(28, 'SEQ x \\* ALPHABETIC')).toBe('BB')
  })
})
