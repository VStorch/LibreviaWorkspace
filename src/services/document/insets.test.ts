import { describe, expect, it } from 'vitest'
import { contentHeightMm, contentInsetsMm, DEFAULT_PAGE_SETUP, PageSize, type PageSetup } from './model.js'

/** A4 with the corpus margins and bands 12.5 mm from the edge. */
const page: PageSetup = {
  ...DEFAULT_PAGE_SETUP,
  size: PageSize.A4,
  margins: { top: 25, right: 30, bottom: 25, left: 30 },
  headerDistanceMm: 12.5,
  footerDistanceMm: 12.5,
}

describe('onde a coluna de texto começa e termina', () => {
  it('faixa que cabe na margem não desloca nada', () => {
    // 12.5 + 8 = 20.5 mm, inside the 25 mm margin: the body stays where it was.
    const inset = contentInsetsMm(page, { headerMm: 8, footerMm: 8 })

    expect(inset.top).toBe(25)
    expect(inset.bottom).toBe(25)
  })

  it('cabeçalho mais alto que a margem empurra o corpo para baixo', () => {
    // The corporate grid header, four rows and a logo, exceeds the margin.
    const inset = contentInsetsMm(page, { headerMm: 23, footerMm: 0 })

    expect(inset.top).toBe(35.5)
    expect(inset.bottom).toBe(25)
  })

  it('o rodapé sobe o fim da coluna pela mesma conta', () => {
    const inset = contentInsetsMm(page, { headerMm: 0, footerMm: 20 })

    expect(inset.bottom).toBe(32.5)
  })

  it('a altura útil desconta o que a faixa tomou', () => {
    // 297 − 35.5 − 25: this height decides where the sheet breaks, and it must be the same on
    // screen and paper.
    expect(contentHeightMm(page, { headerMm: 23, footerMm: 0 })).toBeCloseTo(236.5, 1)
    expect(contentHeightMm(page)).toBeCloseTo(247, 1)
  })
})
