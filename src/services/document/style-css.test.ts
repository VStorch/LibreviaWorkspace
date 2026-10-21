import { describe, expect, it } from 'vitest'
import { styleSheetCss } from './style-css.js'
import { BUILTIN_STYLES, LEGACY_STYLES } from './styles.js'

function ruleOf(css: string, selector: string): string {
  const line = css.split('\n').find((candidate) => candidate.startsWith(`${selector} {`))
  if (line === undefined) throw new Error(`sem regra para ${selector}`)
  return line
}

describe('styleSheetCss', () => {
  it('an old file renders as before styles: Times 12, line spacing 1.5', () => {
    const css = styleSheetCss(LEGACY_STYLES)
    expect(ruleOf(css, '.page__content')).toContain("font-family: 'Times New Roman', serif;")
    expect(ruleOf(css, '.page__content')).toContain('line-height: 1.4997;')
    expect(ruleOf(css, '.page__content > p:not([data-style-id])')).toContain('margin-bottom: 12pt;')
    expect(ruleOf(css, '.page__content > h1:not([data-style-id])')).toContain('font-size: 22pt;')
  })

  it('a new document starts as Calibri 11, line spacing 1.08, 8 pt after', () => {
    const css = styleSheetCss(BUILTIN_STYLES)
    const body = ruleOf(css, '.page__content > p:not([data-style-id])')
    expect(body).toContain("font-family: 'Calibri', sans-serif;")
    expect(body).toContain('font-size: 11pt;')
    expect(body).toContain('margin-top: 0pt;')
    expect(body).toContain('margin-bottom: 8pt;')
    expect(ruleOf(css, '.page__content > h1:not([data-style-id])')).toContain('font-weight: 400;')
  })

  it('every paragraph style gets a rule by id, and a character style does not', () => {
    const css = styleSheetCss(BUILTIN_STYLES)
    expect(ruleOf(css, '.page__content > [data-style-id="ListParagraph"]')).toContain('padding-left: 12.7mm;')
    expect(css).not.toContain('[data-style-id="Hyperlink"]')
  })

  it('a heading the document does not define is drawn as the writer will save it', () => {
    const styles = Object.fromEntries(
      Object.entries(BUILTIN_STYLES.styles).filter(([id]) => id !== 'Heading2'),
    )
    const css = styleSheetCss({ ...BUILTIN_STYLES, styles })
    expect(ruleOf(css, '.page__content > h2:not([data-style-id])')).toContain('font-size: 17pt;')
  })

  it('an id the document does not define gets only the defaults, before the rules by id', () => {
    const css = styleSheetCss(BUILTIN_STYLES)
    const generic = ruleOf(css, '.page__content > [data-style-id]')
    expect(generic).toContain('margin-bottom: 8pt;')
    expect(css.indexOf('.page__content > [data-style-id] {')).toBeLessThan(
      css.indexOf('.page__content > [data-style-id="Normal"]'),
    )
  })

  it('the black declared in the style becomes a variable that only the dark theme redefines', () => {
    const normal = BUILTIN_STYLES.styles['Normal']!
    const styles = {
      ...BUILTIN_STYLES.styles,
      Normal: { ...normal, character: { ...normal.character, color: '#000000' } },
    }
    const css = styleSheetCss({ ...BUILTIN_STYLES, styles })
    expect(ruleOf(css, '.page__content > [data-style-id="Normal"]')).toContain(
      'color: var(--declared-black, #000000);',
    )
  })

  it("a color that is not black stays the author's", () => {
    const css = styleSheetCss(BUILTIN_STYLES)
    expect(ruleOf(css, '.page__content > [data-style-id="Heading1"]')).toContain('color: #2f5496;')
  })
})
