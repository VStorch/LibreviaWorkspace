import { describe, expect, it } from 'vitest'
import { DOCUMENT_FONT_CSS } from './fonts.js'

describe('@font-face rules', () => {
  it('declares the five families by the name the document uses', () => {
    // The declared name is the **original** font's, not the substitute's: that is what lets
    // `w:rFonts w:ascii="Calibri"` find Carlito without rewriting the document.
    for (const family of ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New']) {
      expect(DOCUMENT_FONT_CSS).toContain(`font-family: '${family}'`)
    }
  })

  it('prefers the font installed on the machine over the bundled one', () => {
    // `local()` before `url()`: whoever has the real Calibri deserves the real Calibri, and it
    // saves loading a file.
    const firstRule = DOCUMENT_FONT_CSS.slice(0, DOCUMENT_FONT_CSS.indexOf('}'))
    expect(firstRule.indexOf('local(')).toBeLessThan(firstRule.indexOf('url('))
  })

  it('local() of each style names the style, not the family', () => {
    // `local()` matches by font name: `local('Liberation Sans')` in the bold rule would serve the
    // regular face as bold where Liberation is installed.
    const bold = DOCUMENT_FONT_CSS.split('@font-face').find(
      (rule) => rule.includes("font-family: 'Arial'") && rule.includes('font-weight: 700'),
    )

    expect(bold).toContain("local('Liberation Sans Bold')")
    expect(bold).toContain("local('LiberationSans-Bold')")
    expect(bold).not.toContain("local('Liberation Sans')")
  })

  it('does not use a backtick, which would end the CSS template literal', () => {
    // `DOCUMENT_CONTENT_CSS` interpolates this inside a template literal. A backtick here does not
    // break this file; it breaks another one's compilation, with a syntax error that does not point
    // here.
    expect(DOCUMENT_FONT_CSS).not.toContain('`')
  })
})
