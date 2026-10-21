import { describe, expect, it } from 'vitest'
import {
  GUARANTEED_FONT_FAMILIES,
  familiesInDocument,
  firstFamilyOf,
  orderFontFamilies,
  parseFontconfigFamilies,
  parseWindowsFontRegistry,
} from './font-list.js'

describe('font list', () => {
  it("the document's fonts come first", () => {
    // Someone opening another person's file looks for its font, not among three hundred.
    const ordered = orderFontFamilies(['Arial', 'Zapfino'], ['Garamond'])

    expect(ordered[0]).toBe('Garamond')
  })

  it('guaranteed fonts come before merely installed ones', () => {
    const ordered = orderFontFamilies(['AAA Fonte', 'Calibri'])

    expect(ordered.slice(0, GUARANTEED_FONT_FAMILIES.length)).toEqual([...GUARANTEED_FONT_FAMILIES])
    expect(ordered.at(-1)).toBe('AAA Fonte')
  })

  it('an empty system list does not empty the picker', () => {
    // A system without `fontconfig` is not an error: the toolbar offers what the installer
    // guarantees.
    expect(orderFontFamilies([])).toEqual([...GUARANTEED_FONT_FAMILIES])
  })

  it('the same name does not appear twice, even in another case', () => {
    const ordered = orderFontFamilies(['arial', 'ARIAL', 'Arial'], ['Arial'])

    expect(ordered.filter((family) => family.toLowerCase() === 'arial')).toHaveLength(1)
  })

  it('installed fonts come out in alphabetical order', () => {
    // `fc-list` returns fontconfig cache order, which is no order at all to a reader.
    const ordered = orderFontFamilies(['Ubuntu', 'DejaVu Sans', 'Noto Serif'])
    const extras = ordered.filter((family) => !GUARANTEED_FONT_FAMILIES.includes(family))

    expect(extras).toEqual(['DejaVu Sans', 'Noto Serif', 'Ubuntu'])
  })

  it("the document's CSS stack becomes a name", () => {
    expect(firstFamilyOf("'Times New Roman', Liberation Serif, serif")).toBe('Times New Roman')
    expect(firstFamilyOf('')).toBe('')
  })

  it('finds the font in the text mark and in the block attribute', () => {
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

  it('reads fontconfig output', () => {
    const output = 'DejaVu Sans\nNimbus Sans,Nimbus Sans L\n\nUbuntu\nDejaVu Sans\n'

    expect(parseFontconfigFamilies(output)).toEqual(['DejaVu Sans', 'Nimbus Sans', 'Ubuntu'])
  })

  it('reads the Windows registry without repeating a style as a family', () => {
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
