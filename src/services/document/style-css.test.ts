import { describe, expect, it } from 'vitest'
import { styleSheetCss } from './style-css.js'
import { BUILTIN_STYLES, LEGACY_STYLES } from './styles.js'

function ruleOf(css: string, selector: string): string {
  const line = css.split('\n').find((candidate) => candidate.startsWith(`${selector} {`))
  if (line === undefined) throw new Error(`sem regra para ${selector}`)
  return line
}

describe('styleSheetCss', () => {
  it('o arquivo antigo desenha como antes dos estilos: Times 12, entrelinha 1,5', () => {
    const css = styleSheetCss(LEGACY_STYLES)
    expect(ruleOf(css, '.page__content')).toContain("font-family: 'Times New Roman', serif;")
    expect(ruleOf(css, '.page__content')).toContain('line-height: 1.4997;')
    expect(ruleOf(css, '.page__content > p:not([data-style-id])')).toContain('margin-bottom: 12pt;')
    expect(ruleOf(css, '.page__content > h1:not([data-style-id])')).toContain('font-size: 22pt;')
  })

  it('o documento novo nasce Calibri 11, entrelinha 1,08, 8 pt depois', () => {
    const css = styleSheetCss(BUILTIN_STYLES)
    const body = ruleOf(css, '.page__content > p:not([data-style-id])')
    expect(body).toContain("font-family: 'Calibri', sans-serif;")
    expect(body).toContain('font-size: 11pt;')
    expect(body).toContain('margin-top: 0pt;')
    expect(body).toContain('margin-bottom: 8pt;')
    expect(ruleOf(css, '.page__content > h1:not([data-style-id])')).toContain('font-weight: 400;')
  })

  it('todo estilo de parágrafo ganha regra pelo id, e o de caractere não', () => {
    const css = styleSheetCss(BUILTIN_STYLES)
    expect(ruleOf(css, '.page__content > [data-style-id="ListParagraph"]')).toContain('padding-left: 12.7mm;')
    expect(css).not.toContain('[data-style-id="Hyperlink"]')
  })

  it('título que o documento não define é desenhado como o escritor vai gravá-lo', () => {
    const styles = Object.fromEntries(
      Object.entries(BUILTIN_STYLES.styles).filter(([id]) => id !== 'Heading2'),
    )
    const css = styleSheetCss({ ...BUILTIN_STYLES, styles })
    expect(ruleOf(css, '.page__content > h2:not([data-style-id])')).toContain('font-size: 17pt;')
  })
})
