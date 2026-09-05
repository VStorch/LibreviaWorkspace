import { describe, expect, it } from 'vitest'
import { bandFloatsOf, frameOf, placeFloating, type FloatingObject } from './floating.js'
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

  it('a peça de um grupo soma o deslocamento dela depois de resolver a âncora', () => {
    // Um cabeçalho corporativo é um grupo de formas: a âncora diz onde o grupo
    // está, e cada peça tem a coordenada dela dentro dele. Somar depois é o que
    // faz a conta valer também quando a âncora traz alinhamento em vez de
    // deslocamento — aí quem resolve a origem é esta função, e não o arquivo.
    const box = placeFloating(object({ hFrom: 'page', hOffsetMm: 31.75, dxMm: 129, dyMm: 6.9 }), page, 0)

    expect(box.leftMm).toBeCloseTo(160.75, 2)
    expect(box.topMm).toBeCloseTo(6.9, 2)
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

describe('objetos ancorados das faixas', () => {
  const box: FloatingObject = {
    kind: 'text',
    widthMm: 84.8,
    heightMm: 9.1,
    rotation: 0,
    hFrom: 'page',
    hOffsetMm: 31.75,
    vFrom: 'page',
    vOffsetMm: 2.7,
    behind: false,
    wrap: 'square',
    content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Folha {n}' }] }],
  }

  const withBand: PageSetup = {
    ...page,
    headerBand: { left: [], center: [], right: [], rule: false, rows: [], floats: [box] },
  }

  it('a caixa do cabeçalho recebe o número da folha em que está', () => {
    // O campo `PAGE` do Word chega como `{n}` — o mesmo marcador do cabeçalho
    // digitado à mão. Sem a troca, a folha sairia com as chaves escritas nela.
    const terceira = bandFloatsOf(withBand, 3)[0]!
    const texto = terceira.object.content![0]!.content![0]!.text

    expect(texto).toBe('Folha 3')
  })

  it('a substituição não altera o objeto original', () => {
    bandFloatsOf(withBand, 3)
    expect(box.content![0]!.content![0]!.text).toBe('Folha {n}')
  })
})

/**
 * A moldura da forma em CSS.
 *
 * O aviso "moldura e preenchimento de formas" saía em toda caixa de texto,
 * tivesse ela decoração ou não. Agora a decoração é desenhada, e o aviso fala
 * só do que sobra.
 */
describe('frameOf', () => {
  const caixa = (extra: Partial<FloatingObject>): FloatingObject => ({
    kind: 'text',
    widthMm: 80,
    heightMm: 10,
    rotation: 0,
    hFrom: 'page',
    vFrom: 'paragraph',
    behind: false,
    wrap: 'none',
    ...extra,
  })

  it('desenha o preenchimento e o traço que o documento declara', () => {
    expect(frameOf(caixa({ fill: '#ffffff', line: '#1f5fa9', lineWidthPt: 1 }))).toEqual({
      background: '#ffffff',
      border: '1pt solid #1f5fa9',
    })
  })

  it('traço de espessura zero não é traço', () => {
    // As caixas do cabeçalho do corpus declaram cor nenhuma e espessura zero:
    // não há moldura, e desenhar uma poria uma borda que o documento não tem.
    expect(frameOf(caixa({ lineWidthPt: 0 }))).toEqual({})
    expect(frameOf(caixa({ line: '#000000', lineWidthPt: 0 }))).toEqual({})
  })

  it('o tracejado sai tracejado', () => {
    expect(frameOf(caixa({ line: '#000000', lineWidthPt: 0.75, dash: true }))).toEqual({
      border: '0.75pt dashed #000000',
    })
  })
})
