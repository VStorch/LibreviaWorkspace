import { describe, expect, it } from 'vitest'
import { placeFloating, type FloatingObject } from './floating.js'
import { DEFAULT_PAGE_SETUP, PageSize, type PageSetup } from './model.js'

/** A4 retrato com 30 mm nas laterais: a geometria da capa do corpus. */
const page: PageSetup = {
  ...DEFAULT_PAGE_SETUP,
  size: PageSize.A4,
  margins: { top: 35, right: 30, bottom: 25, left: 30 },
}

function object(overrides: Partial<FloatingObject> = {}): FloatingObject {
  return {
    kind: 'image',
    widthMm: 50,
    heightMm: 20,
    rotation: 0,
    hFrom: 'column',
    hOffsetMm: 0,
    vFrom: 'paragraph',
    vOffsetMm: 0,
    behind: false,
    wrap: 'none',
    ...overrides,
  }
}

describe('posição de objeto ancorado', () => {
  it('deslocamento a partir da coluna começa na margem esquerda', () => {
    const box = placeFloating(object({ hFrom: 'column', hOffsetMm: 10 }), page, 0)
    expect(box.leftMm).toBe(40)
  })

  it('deslocamento a partir da página começa na borda do papel', () => {
    const box = placeFloating(object({ hFrom: 'page', hOffsetMm: 10 }), page, 0)
    expect(box.leftMm).toBe(10)
  })

  it('deslocamento negativo leva o objeto para fora da coluna', () => {
    // É assim que a marca da capa sai para a margem: 126,65 mm à esquerda do
    // começo da coluna. Tratar o valor como sem sinal a jogaria para o meio da
    // página, sobre o texto.
    const box = placeFloating(object({ hFrom: 'column', hOffsetMm: -126.65 }), page, 0)
    expect(box.leftMm).toBeCloseTo(-96.65, 2)
  })

  it('centralizado usa a coluna de texto, não a folha', () => {
    // Coluna de 150 mm começando em 30: um objeto de 50 mm centra em 80.
    const box = placeFloating(object({ hAlign: 'center', hOffsetMm: undefined }), page, 0)
    expect(box.leftMm).toBe(80)
  })

  it('alinhado à direita encosta no fim da coluna', () => {
    const box = placeFloating(object({ hAlign: 'right', hOffsetMm: undefined }), page, 0)
    expect(box.leftMm).toBe(130)
  })

  it('vertical a partir do parágrafo conta da altura dele na folha', () => {
    // É a origem mais comum, e a única que este módulo não resolve sozinho: o
    // parágrafo só tem posição depois de paginar.
    const box = placeFloating(object({ vFrom: 'paragraph', vOffsetMm: 58.19 }), page, 40)
    expect(box.topMm).toBeCloseTo(98.19, 2)
  })

  it('vertical a partir da margem ignora onde o parágrafo caiu', () => {
    const box = placeFloating(object({ vFrom: 'margin', vOffsetMm: 5 }), page, 999)
    expect(box.topMm).toBe(40)
  })

  it('vertical a partir da página começa na borda de cima', () => {
    const box = placeFloating(object({ vFrom: 'page', vOffsetMm: 5 }), page, 999)
    expect(box.topMm).toBe(5)
  })

  it('a rotação não mexe nas medidas', () => {
    // O Word posiciona a caixa sem girar e depois a gira em torno do centro,
    // que é o que `transform: rotate()` faz. Trocar largura por altura aqui
    // deslocaria o objeto por metade da diferença entre as duas.
    const box = placeFloating(object({ rotation: 270, widthMm: 285.76, heightMm: 80.14 }), page, 0)

    expect(box.widthMm).toBe(285.76)
    expect(box.heightMm).toBe(80.14)
    expect(box.rotation).toBe(270)
  })

  it('a marca da capa cai na lateral esquerda, e não sobre o texto', () => {
    // O caso completo do corpus, com os números do arquivo. Depois de girada em
    // torno do centro, uma caixa de 285,76 × 80,14 mm posta em −96,65 mm ocupa
    // de 6,2 a 86,3 mm — a faixa lateral, à esquerda da coluna de texto.
    const marca = object({
      rotation: 270,
      widthMm: 285.76,
      heightMm: 80.14,
      hFrom: 'column',
      hOffsetMm: -126.65,
      vFrom: 'paragraph',
      vOffsetMm: 58.19,
      behind: true,
    })

    const box = placeFloating(marca, page, 45)
    const centro = box.leftMm + box.widthMm / 2
    const esquerdaGirada = centro - box.heightMm / 2

    expect(esquerdaGirada).toBeCloseTo(6.2, 1)
    expect(esquerdaGirada + box.heightMm).toBeLessThan(page.margins.left + 60)
  })
})
