import { describe, expect, it } from 'vitest'
import { contentHeightMm, contentInsetsMm, DEFAULT_PAGE_SETUP, PageSize, type PageSetup } from './model.js'

/** A4 com as margens do corpus e as faixas a 12,5 mm da borda. */
const page: PageSetup = {
  ...DEFAULT_PAGE_SETUP,
  size: PageSize.A4,
  margins: { top: 25, right: 30, bottom: 25, left: 30 },
  headerDistanceMm: 12.5,
  footerDistanceMm: 12.5,
}

describe('onde a coluna de texto começa e termina', () => {
  it('faixa que cabe na margem não desloca nada', () => {
    // 12,5 + 8 = 20,5 mm, dentro dos 25 da margem: o corpo fica onde estava.
    const inset = contentInsetsMm(page, { headerMm: 8, footerMm: 8 })

    expect(inset.top).toBe(25)
    expect(inset.bottom).toBe(25)
  })

  it('cabeçalho mais alto que a margem empurra o corpo para baixo', () => {
    // O caso do cabeçalho corporativo em grade: quatro linhas e um logotipo.
    // Sem isto, a primeira linha do texto era escrita por cima da última do
    // cabeçalho — as duas se encontravam por oito milímetros no corpus.
    const inset = contentInsetsMm(page, { headerMm: 23, footerMm: 0 })

    expect(inset.top).toBe(35.5)
    expect(inset.bottom).toBe(25)
  })

  it('o rodapé sobe o fim da coluna pela mesma conta', () => {
    const inset = contentInsetsMm(page, { headerMm: 0, footerMm: 20 })

    expect(inset.bottom).toBe(32.5)
  })

  it('a altura útil desconta o que a faixa tomou', () => {
    // 297 − 35,5 − 25: é essa altura que decide onde a folha quebra, e ela tem
    // de ser a mesma na tela e no papel.
    expect(contentHeightMm(page, { headerMm: 23, footerMm: 0 })).toBeCloseTo(236.5, 1)
    expect(contentHeightMm(page)).toBeCloseTo(247, 1)
  })
})
