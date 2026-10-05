import { describe, expect, it } from 'vitest'
import {
  GUARANTEED_FONT_FAMILIES,
  familiesInDocument,
  firstFamilyOf,
  orderFontFamilies,
  parseFontconfigFamilies,
  parseWindowsFontRegistry,
} from './font-list.js'

describe('lista de fontes', () => {
  it('as do documento vêm primeiro', () => {
    // Someone opening another person's file looks for its font, not among three hundred.
    const ordered = orderFontFamilies(['Arial', 'Zapfino'], ['Garamond'])

    expect(ordered[0]).toBe('Garamond')
  })

  it('as garantidas vêm antes das apenas instaladas', () => {
    const ordered = orderFontFamilies(['AAA Fonte', 'Calibri'])

    expect(ordered.slice(0, GUARANTEED_FONT_FAMILIES.length)).toEqual([...GUARANTEED_FONT_FAMILIES])
    expect(ordered.at(-1)).toBe('AAA Fonte')
  })

  it('lista vazia do sistema não esvazia o seletor', () => {
    // A system without `fontconfig` is not an error: the toolbar offers what the installer
    // guarantees.
    expect(orderFontFamilies([])).toEqual([...GUARANTEED_FONT_FAMILIES])
  })

  it('o mesmo nome não aparece duas vezes, nem com outra caixa', () => {
    const ordered = orderFontFamilies(['arial', 'ARIAL', 'Arial'], ['Arial'])

    expect(ordered.filter((family) => family.toLowerCase() === 'arial')).toHaveLength(1)
  })

  it('as instaladas saem em ordem alfabética', () => {
    // `fc-list` returns fontconfig cache order, which is no order at all to a reader.
    const ordered = orderFontFamilies(['Ubuntu', 'DejaVu Sans', 'Noto Serif'])
    const extras = ordered.filter((family) => !GUARANTEED_FONT_FAMILIES.includes(family))

    expect(extras).toEqual(['DejaVu Sans', 'Noto Serif', 'Ubuntu'])
  })

  it('a pilha de CSS do documento vira um nome', () => {
    expect(firstFamilyOf("'Times New Roman', Liberation Serif, serif")).toBe('Times New Roman')
    expect(firstFamilyOf('')).toBe('')
  })

  it('acha a fonte na marca do texto e no atributo do bloco', () => {
    // The reader emits the font in both places: line height comes from the element's font, not from
    // what is written inside it.
    const doc = {
      type: 'doc',
      content: [
        {
          type: 'heading',
          attrs: { fontFamily: 'Cambria, Caladea, serif' },
          content: [{ type: 'text', marks: [{ type: 'textStyle', attrs: { fontFamily: 'Garamond' } }] }],
        },
      ],
    }

    expect(familiesInDocument(doc)).toEqual(['Cambria', 'Garamond'])
  })

  it('lê a saída do fontconfig', () => {
    const output = 'DejaVu Sans\nNimbus Sans,Nimbus Sans L\n\nUbuntu\nDejaVu Sans\n'

    expect(parseFontconfigFamilies(output)).toEqual(['DejaVu Sans', 'Nimbus Sans', 'Ubuntu'])
  })

  it('lê o registro do Windows sem repetir corte como família', () => {
    // The registry stores one file per style. Without cutting the suffix, the picker would offer
    // "Arial", "Arial Bold" and "Arial Italic" as three fonts.
    const output = [
      'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts',
      '    Arial & Arial Bold & Arial Italic (TrueType)    REG_SZ    arial.ttf',
      '    Arial Black (TrueType)    REG_SZ    ariblk.ttf',
      '    Segoe UI Light (TrueType)    REG_SZ    segoeuil.ttf',
      '',
    ].join('\r\n')

    expect(parseWindowsFontRegistry(output)).toEqual(['Arial', 'Arial Black', 'Segoe UI Light'])
  })
})
