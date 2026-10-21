import { describe, expect, it } from 'vitest'
import {
  DEFAULT_NATURAL_LINE_HEIGHT,
  cssLineHeightOf,
  firstFontOf,
  lineFactorOf,
  naturalLineHeightOf,
} from './line-metrics.js'

describe('natural line height', () => {
  it('the CSS stack is read by its first family', () => {
    expect(firstFontOf('Calibri, sans-serif')).toBe('Calibri')
    expect(firstFontOf('"Times New Roman", serif')).toBe('Times New Roman')
    expect(firstFontOf(null)).toBeNull()
    expect(naturalLineHeightOf('Calibri, sans-serif')).toBe(1.2207)
  })

  it("with no declared font the editor's applies; an outside font has no guess", () => {
    // Both answers come from `LineMetrics.Of`: silence is our font, and a font the installer does
    // not ship falls back to a machine-dependent substitute.
    expect(naturalLineHeightOf(undefined)).toBe(DEFAULT_NATURAL_LINE_HEIGHT)
    expect(naturalLineHeightOf('Aptos')).toBeNull()
  })

  it('the Word factor and the CSS number round-trip in both directions', () => {
    // 1.5 lines in Calibri is 1.83105 in CSS. What the dialog picks must come back the same, or
    // opening and closing the dialog changes the document.
    expect(cssLineHeightOf(1.5, 'Calibri, sans-serif')).toBe('1.8311')
    expect(lineFactorOf(1.8311, 'Calibri, sans-serif')).toBe(1.5)

    expect(cssLineHeightOf(1, 'Arial')).toBe('1.1499')
    expect(lineFactorOf(1.1499, 'Arial')).toBe(1)
  })

  it('unknown font with factor 1 comes out as "normal"', () => {
    // That is what the reader does: without knowing the height, the browser measures best.
    expect(cssLineHeightOf(1, 'Aptos')).toBe('normal')
    expect(cssLineHeightOf(2, 'Aptos')).toBe('2.2998')
    expect(lineFactorOf(2.2998, 'Aptos')).toBe(2)
  })
})
