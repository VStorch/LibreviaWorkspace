import { describe, expect, it } from 'vitest'
import { DOCUMENT_FONT_CSS } from './fonts.js'

describe('regras @font-face', () => {
  it('declara as cinco famílias pelo nome que o documento usa', () => {
    // O nome declarado é o da fonte **original**, não o da substituta: é o que
    // faz um `w:rFonts w:ascii="Calibri"` achar a Carlito sem que ninguém
    // reescreva o documento.
    for (const familia of ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New']) {
      expect(DOCUMENT_FONT_CSS).toContain(`font-family: '${familia}'`)
    }
  })

  it('prefere a fonte instalada na máquina à empacotada', () => {
    // `local()` antes de `url()`: quem tem a Calibri de verdade merece a
    // Calibri de verdade, e ainda poupa carregar arquivo à toa.
    const primeiraRegra = DOCUMENT_FONT_CSS.slice(0, DOCUMENT_FONT_CSS.indexOf('}'))
    expect(primeiraRegra.indexOf('local(')).toBeLessThan(primeiraRegra.indexOf('url('))
  })

  it('o local() de cada corte nomeia o corte, e não a família', () => {
    // `local()` casa por nome de fonte, não de família: `local('Liberation
    // Sans')` dentro da regra de negrito acha a normal e a serve como se fosse
    // negrito. Numa máquina com as Liberation instaladas — todo Linux de
    // escritório — era isso que apagava o negrito de todo documento importado.
    const negrito = DOCUMENT_FONT_CSS.split('@font-face').find(
      (regra) => regra.includes("font-family: 'Arial'") && regra.includes('font-weight: 700'),
    )

    expect(negrito).toContain("local('Liberation Sans Bold')")
    expect(negrito).toContain("local('LiberationSans-Bold')")
    expect(negrito).not.toContain("local('Liberation Sans')")
  })

  it('não usa crase, que encerraria o template literal do CSS', () => {
    // `DOCUMENT_CONTENT_CSS` interpola isto dentro de um template literal. Uma
    // crase aqui não quebra este arquivo — quebra a compilação de outro, com
    // erro de sintaxe que não aponta para cá.
    expect(DOCUMENT_FONT_CSS).not.toContain('`')
  })
})
