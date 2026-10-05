import { describe, expect, it } from 'vitest'
import { DOCUMENT_FONT_CSS } from './fonts.js'

describe('regras @font-face', () => {
  it('declara as cinco famílias pelo nome que o documento usa', () => {
    // The declared name is the **original** font's, not the substitute's: that is what lets
    // `w:rFonts w:ascii="Calibri"` find Carlito without rewriting the document.
    for (const familia of ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New']) {
      expect(DOCUMENT_FONT_CSS).toContain(`font-family: '${familia}'`)
    }
  })

  it('prefere a fonte instalada na máquina à empacotada', () => {
    // `local()` before `url()`: whoever has the real Calibri deserves the real Calibri, and it
    // saves loading a file.
    const primeiraRegra = DOCUMENT_FONT_CSS.slice(0, DOCUMENT_FONT_CSS.indexOf('}'))
    expect(primeiraRegra.indexOf('local(')).toBeLessThan(primeiraRegra.indexOf('url('))
  })

  it('o local() de cada corte nomeia o corte, e não a família', () => {
    // `local()` matches by font name: `local('Liberation Sans')` in the bold rule would serve the
    // regular face as bold where Liberation is installed.
    const negrito = DOCUMENT_FONT_CSS.split('@font-face').find(
      (regra) => regra.includes("font-family: 'Arial'") && regra.includes('font-weight: 700'),
    )

    expect(negrito).toContain("local('Liberation Sans Bold')")
    expect(negrito).toContain("local('LiberationSans-Bold')")
    expect(negrito).not.toContain("local('Liberation Sans')")
  })

  it('não usa crase, que encerraria o template literal do CSS', () => {
    // `DOCUMENT_CONTENT_CSS` interpolates this inside a template literal. A backtick here does not
    // break this file; it breaks another one's compilation, with a syntax error that does not point
    // here.
    expect(DOCUMENT_FONT_CSS).not.toContain('`')
  })
})
