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

  it('não usa crase, que encerraria o template literal do CSS', () => {
    // `DOCUMENT_CONTENT_CSS` interpola isto dentro de um template literal. Uma
    // crase aqui não quebra este arquivo — quebra a compilação de outro, com
    // erro de sintaxe que não aponta para cá.
    expect(DOCUMENT_FONT_CSS).not.toContain('`')
  })
})
