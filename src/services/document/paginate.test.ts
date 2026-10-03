import { describe, expect, it } from 'vitest'
import {
  paginate,
  paginateSections,
  type MeasuredBlock,
  type MeasuredNote,
  type SectionFlow,
  noteLineTop,
  noteSpan,
} from './paginate.js'

/**
 * Blocos empilhados de altura fixa, na ordem — o formato que o editor mede.
 */
function stack(
  heights: readonly number[],
  marks: { pageBreak?: number[]; breakAfter?: number[]; keepNext?: number[] } = {},
): MeasuredBlock[] {
  let top = 0
  return heights.map((height, index) => {
    const block: MeasuredBlock = {
      top,
      height,
      breakpoints: [],
      isPageBreak: marks.pageBreak?.includes(index) ?? false,
      breakAfter: marks.breakAfter?.includes(index) ?? false,
      keepWithNext: marks.keepNext?.includes(index) ?? false,
    }
    top += height
    return block
  })
}

describe('paginação', () => {
  it('documento que cabe numa folha não quebra', () => {
    expect(paginate(stack([100, 100, 100]), 1000)).toEqual([])
  })

  it('quebra antes do bloco que estouraria a página', () => {
    // Três blocos de 400 numa página de 1000: o terceiro não cabe, e a quebra
    // cai no topo dele — nunca no meio.
    expect(paginate(stack([400, 400, 400]), 1000)).toEqual([800])
  })

  it('a quebra pedida à mão vale com a página pela metade', () => {
    // O medidor anterior ignorava o nó `pageBreak`: num documento com capa e
    // sumário, as três páginas apareciam como uma só.
    const blocos = stack([100, 10, 100], { pageBreak: [1] })
    expect(paginate(blocos, 1000)).toEqual([110])
  })

  it('quebra à mão no começo do documento não cria página vazia', () => {
    // Uma quebra antes de qualquer conteúdo não tem página para fechar.
    const blocos = stack([10, 100], { pageBreak: [0] })
    expect(paginate(blocos, 1000)).toEqual([10])
  })

  it('título não fica sozinho no pé da página', () => {
    // O bloco 2 é um título que estoura junto com o parágrafo dele: os dois
    // descem. É o `break-after: avoid` que a exportação aplica, e sem isto a
    // marca da tela cai um bloco depois de onde o PDF quebra.
    const blocos = stack([600, 200, 100, 300], { keepNext: [2] })
    expect(paginate(blocos, 1000)).toEqual([800])
  })

  it('um título não arrasta a página inteira atrás de si', () => {
    // Todos os três pedem para ficar com o seguinte. O terceiro estoura, puxa o
    // segundo junto — e a corrente para no primeiro, que já está no topo da
    // página: descer todo mundo deixaria a folha em branco.
    const blocos = stack([400, 400, 400], { keepNext: [0, 1, 2] })
    expect(paginate(blocos, 1000)).toEqual([400])
  })

  it('bloco mais alto que a página fica com a folha só para si', () => {
    // Uma captura de tela grande não tem onde ser cortada: ela abre folha nova,
    // transborda, e o que vem depois abre outra. Sem o segundo corte o
    // parágrafo seguinte encostaria embaixo do transbordo, fora do papel.
    const blocos = stack([100, 3000, 100])
    expect(paginate(blocos, 1000)).toEqual([100, 3100])
  })

  it('bloco gigante no fim do documento não cria folha em branco', () => {
    // O contrapeso: sem nada depois dele, o segundo corte abriria uma página
    // vazia no fim.
    const blocos = stack([100, 3000])
    expect(paginate(blocos, 1000)).toEqual([100])
  })

  it('a quebra que o parágrafo carrega termina a folha depois dele', () => {
    // `w:br w:type="page"` dentro de um `w:r`: o Word grava assim quando a
    // quebra encerra o parágrafo. Vira propriedade do bloco, porque um nó de
    // bloco em posição de linha é inválido no editor.
    const blocos = stack([100, 100, 100], { breakAfter: [1] })
    expect(paginate(blocos, 1000)).toEqual([200])
  })

  it('quebra carregada pelo último bloco não cria folha em branco', () => {
    const blocos = stack([100, 100], { breakAfter: [1] })
    expect(paginate(blocos, 1000)).toEqual([])
  })

  it('altura de página inválida não quebra nada', () => {
    // Margens que somam mais que o papel produzem altura negativa. Sem esta
    // guarda o laço rodaria até o teto de páginas a cada tecla digitada.
    expect(paginate(stack([100, 100]), 0)).toEqual([])
    expect(paginate(stack([100, 100]), -50)).toEqual([])
  })

  it('documento longo produz uma quebra por página cheia', () => {
    const blocos = stack(Array.from({ length: 20 }, () => 250))
    // 250 × 4 = 1000 por página; 20 blocos dão cinco páginas, quatro cortes.
    expect(paginate(blocos, 1000)).toEqual([1000, 2000, 3000, 4000])
  })

  it('documento com mais de quinhentas folhas não empilha o resto na última', () => {
    // Havia um teto de quinhentas páginas: alcançado, o laço parava e todo o
    // resto do documento ficava amontoado na última folha, fora da vista. Uma
    // altura de página pequena — margens absurdas, fonte que não carregou — o
    // alcançava num documento comum.
    const blocos = stack(Array.from({ length: 700 }, () => 100))
    const cortes = paginate(blocos, 100)

    expect(cortes).toHaveLength(699)
    expect(cortes.at(-1)).toBe(69_900)
  })

  it('bloco mais alto que a folha fica com ela só para si, mesmo aos milhares', () => {
    // O outro caminho pelo qual o laço avança. Se ele não avançasse, a paginação
    // travaria a cada tecla digitada — é por isso que o teto existia.
    const blocos = stack(Array.from({ length: 600 }, () => 300))
    expect(paginate(blocos, 100)).toHaveLength(599)
  })
})

describe('cortes dentro de blocos', () => {
  function splittable(
    height: number,
    breakpoints: number[],
    marks: Partial<MeasuredBlock> = {},
  ): MeasuredBlock {
    return { ...stack([height])[0]!, breakpoints, ...marks }
  }

  it('tabela corta na última linha que cabe', () => {
    expect(paginate([splittable(1500, [300, 600, 900, 1200])], 1000)).toEqual([900])
  })

  it('o cabeçalho repetido ocupa a folha seguinte', () => {
    // Linhas de 100, cabeçalho de 100: a segunda folha abre com o cabeçalho e
    // cabe só mais nove linhas — o corte seguinte vem 100 antes.
    const rows = Array.from({ length: 24 }, (_, index) => (index + 1) * 100)
    expect(paginate([splittable(2500, rows, { repeatHeight: 100 })], 1000)).toEqual([1000, 1900])
  })

  it('cabeçalho maior que meia folha não se repete', () => {
    const rows = Array.from({ length: 24 }, (_, index) => (index + 1) * 100)
    expect(paginate([splittable(2500, rows, { repeatHeight: 600 })], 1000)).toEqual([1000, 2000])
  })

  it('tabela de três páginas tem dois cortes internos', () => {
    expect(paginate([splittable(2500, [500, 1000, 1500, 2000])], 1000)).toEqual([1000, 2000])
  })

  it('lista corta no topo do item', () => {
    expect(paginate([splittable(1200, [400, 800])], 1000)).toEqual([800])
  })

  it('bloco atômico gigante continua numa folha só', () => {
    expect(paginate(stack([3000]), 1000)).toEqual([])
  })

  it('quebra explícita precede cortes internos', () => {
    expect(
      paginate(
        [splittable(1500, [500, 1000], { isPageBreak: true }), { ...stack([100])[0]!, top: 1500 }],
        1000,
      ),
    ).toEqual([1500])
  })

  it('quebra depois vale após o último pedaço da tabela', () => {
    expect(
      paginate(
        [splittable(1500, [500, 1000], { breakAfter: true }), { ...stack([100])[0]!, top: 1500 }],
        1000,
      ),
    ).toEqual([1000, 1500])
  })

  it('título acompanha tabela quando nem a primeira linha cabe', () => {
    const blocks = stack([800, 100, 1500], { keepNext: [1] })
    blocks[2] = { ...blocks[2]!, breakpoints: [1200, 1500, 1800, 2100] }
    expect(paginate(blocks, 1000)).toEqual([800, 1800])
  })

  it('keepWithNext não volta para antes de um corte interno', () => {
    const blocks = stack([100, 2500, 100], { keepNext: [0, 1] })
    blocks[1] = { ...blocks[1]!, breakpoints: [600, 1100, 1600, 2100] }
    expect(paginate(blocks, 1000)).toEqual([600, 1600, 2600])
  })

  describe('parágrafo cortado entre linhas', () => {
    // Um parágrafo de dez linhas de 50: os cortes são os topos das linhas 2 a 10.
    const paragraph = (top: number, lines = 10, extra: Partial<MeasuredBlock> = {}): MeasuredBlock => ({
      top,
      height: lines * 50,
      breakpoints: Array.from({ length: lines - 1 }, (_, index) => top + (index + 1) * 50),
      isPageBreak: false,
      breakAfter: false,
      keepWithNext: false,
      ...extra,
    })

    it('a folha termina na última linha que cabe, e não antes do parágrafo', () => {
      // Sem o corte, o parágrafo desceria inteiro e deixaria 300 de buraco na
      // folha.
      expect(paginate([...stack([700]), paragraph(700)], 1000)).toEqual([1000])
    })

    it('a linha que não cabe inteira desce', () => {
      expect(paginate([...stack([720]), paragraph(720)], 1000)).toEqual([970])
    })

    it('manter linhas juntas faz o parágrafo descer inteiro', () => {
      expect(paginate([...stack([700]), paragraph(700, 10, { keepLines: true })], 1000)).toEqual([700])
    })

    it('manter linhas juntas cede quando o parágrafo é maior que a folha', () => {
      expect(paginate([paragraph(0, 30, { keepLines: true })], 1000)).toEqual([1000])
    })

    it('o título fica com as primeiras linhas do parágrafo que ele apresenta', () => {
      const blocks = [...stack([700, 100], { keepNext: [1] }), paragraph(800)]
      expect(paginate(blocks, 1000)).toEqual([1000])
    })

    it('viúvas e órfãs: a última linha não desce sozinha, leva a penúltima', () => {
      // Nove das dez linhas caberiam; a décima ficaria viúva no topo da folha.
      expect(paginate([...stack([550]), paragraph(550, 10, { widowControl: true })], 1000)).toEqual([950])
    })

    it('viúvas e órfãs: a primeira linha não fica sozinha no pé', () => {
      // Só uma linha caberia: o parágrafo inteiro desce.
      expect(paginate([...stack([930]), paragraph(930, 10, { widowControl: true })], 1000)).toEqual([930])
    })

    it('parágrafo de duas ou três linhas anda inteiro', () => {
      expect(paginate([...stack([920]), paragraph(920, 2, { widowControl: true })], 1000)).toEqual([920])
      expect(paginate([...stack([880]), paragraph(880, 3, { widowControl: true })], 1000)).toEqual([880])
    })

    it('sem controle, a linha sozinha é aceita', () => {
      expect(paginate([...stack([930]), paragraph(930, 10, { widowControl: false })], 1000)).toEqual([980])
    })

    it('parágrafo maior que a folha respeita a regra nos dois cortes', () => {
      // 25 linhas numa folha de 20: 20 e 5 não violam, e cortar em 1000 serve.
      expect(paginate([paragraph(0, 25, { widowControl: true })], 1000)).toEqual([1000])
      // 21 linhas: cortar em 20 deixaria a 21ª viúva, então 19 ficam.
      expect(paginate([paragraph(0, 21, { widowControl: true })], 1000)).toEqual([950])
    })

    it('o pé da captura ancorada corta mesmo com o controle de viúvas', () => {
      // Quadro de 900 e a linha vazia dele (50) numa folha de 1000 que já tem
      // 60: a linha desce e o quadro fica, como no LibreOffice.
      const captura: MeasuredBlock = {
        top: 60,
        height: 950,
        breakpoints: [960],
        freeBreakpoints: [960],
        isPageBreak: false,
        breakAfter: false,
        keepWithNext: false,
        widowControl: true,
      }
      expect(paginate([...stack([60]), captura], 1000)).toEqual([960])
    })

    it('a linha vazia da captura sobra no pé da folha em vez de descer', () => {
      // Quadro de 900 + linha de 50 a partir de 60: passa 10 da folha, e cabe,
      // porque a linha vazia entra na margem de baixo. O bloco seguinte abre a
      // folha nova.
      const captura = { ...stack([950])[0]!, top: 60, hangingBottom: 50 }
      const depois = { ...stack([100])[0]!, top: 1010 }
      expect(paginate([...stack([60]), captura, depois], 1000)).toEqual([1010])
    })
  })
})

describe('paginação por seção (M9)', () => {
  const flow = (height: number, options: Partial<SectionFlow> = {}): SectionFlow => ({
    height,
    newSheet: false,
    parity: null,
    restart: null,
    ...options,
  })
  const inSections = (heights: readonly number[], sections: readonly number[]): MeasuredBlock[] =>
    stack(heights).map((block, index) => ({ ...block, section: sections[index] ?? 0 }))

  it('a seção de próxima página corta antes do primeiro bloco dela', () => {
    const plan = paginateSections(inSections([100, 100, 100], [0, 1, 1]), [
      flow(1000),
      flow(1000, { newSheet: true }),
    ])
    expect(plan.breaks).toEqual([100])
    expect(plan.sheets.map((sheet) => [sheet.section, sheet.number, sheet.first])).toEqual([
      [0, 1, true],
      [1, 2, true],
    ])
  })

  it('a contínua continua na mesma folha, e a folha seguinte tem a altura da seção que a abre', () => {
    // A primeira seção tem folha útil de 300; a contínua, de 500 (outras
    // margens). A segunda folha abre na seção 1 e comporta os dois blocos de 250.
    const plan = paginateSections(inSections([200, 50, 250, 250], [0, 1, 1, 1]), [flow(300), flow(500)])
    expect(plan.breaks).toEqual([250])
    expect(plan.sheets.map((sheet) => sheet.section)).toEqual([0, 1])
  })

  it('a seção ímpar que cairia em folha par ganha uma folha em branco antes', () => {
    const plan = paginateSections(inSections([100, 100, 100], [0, 1, 2]), [
      flow(1000),
      flow(1000, { newSheet: true, restart: 1 }),
      flow(1000, { newSheet: true, parity: 'odd' }),
    ])
    expect(plan.breaks).toEqual([100, 200])
    expect(plan.sheets.map((sheet) => [sheet.section, sheet.number, sheet.blank])).toEqual([
      [0, 1, false],
      [1, 1, false],
      [2, 2, true],
      [2, 3, false],
    ])
  })

  it('a seção par que já cai em folha par não ganha folha em branco', () => {
    const plan = paginateSections(inSections([100, 100], [0, 1]), [
      flow(1000),
      flow(1000, { newSheet: true, parity: 'even' }),
    ])
    expect(plan.sheets.map((sheet) => sheet.blank)).toEqual([false, false])
  })

  it('com a folha ainda vazia, a seção nova toma a folha em vez de abrir outra', () => {
    // A quebra de página manual antes da marca de seção deixa a folha nova
    // vazia: o Word não desenha uma segunda folha em branco.
    const blocks = inSections([100, 10, 100], [0, 0, 1]).map((block, index) =>
      index === 1 ? { ...block, isPageBreak: true } : block,
    )
    const plan = paginateSections(blocks, [flow(1000), flow(1000, { newSheet: true, restart: 5 })])
    expect(plan.breaks).toEqual([110])
    expect(plan.sheets.map((sheet) => [sheet.section, sheet.number, sheet.first])).toEqual([
      [0, 1, true],
      [1, 5, true],
    ])
  })
})

describe('colunas (M9)', () => {
  const flow = (height: number, options: Partial<SectionFlow> = {}): SectionFlow => ({
    height,
    newSheet: false,
    parity: null,
    restart: null,
    ...options,
  })
  const inSections = (heights: readonly number[], sections: readonly number[]): MeasuredBlock[] =>
    stack(heights).map((block, index) => ({ ...block, section: sections[index] ?? 0 }))
  const columnsOf = (plan: ReturnType<typeof paginateSections>, count: number): number[] =>
    Array.from({ length: count }, (_, index) => plan.placements.get(index)?.column ?? -1)

  it('enche a primeira coluna, sobe o resto para a segunda e só então abre folha', () => {
    const plan = paginateSections(stack([400, 400, 400, 400, 400]), [flow(1000, { columns: 2 })])
    expect(columnsOf(plan, 5)).toEqual([0, 0, 1, 1, 0])
    // O primeiro da segunda coluna sobe o que a primeira ocupou.
    expect(plan.placements.get(2)?.lift).toBe(-800)
    expect(plan.breaks).toEqual([1600])
    expect(plan.regions.map((region) => [region.sheet, region.top, region.height])).toEqual([
      [0, 0, 800],
      [1, 0, 400],
    ])
  })

  it('antes de uma seção contínua as colunas se equilibram, e o texto seguinte desce ao pé', () => {
    const plan = paginateSections(inSections([100, 100, 100, 100, 50], [0, 0, 0, 0, 1]), [
      flow(1000, { columns: 2 }),
      flow(1000),
    ])
    expect(columnsOf(plan, 4)).toEqual([0, 0, 1, 1])
    expect(plan.regions[0]?.height).toBe(200)
    // O bloco da seção de baixo estava em 400 na tira; desenhado, fica em 200.
    expect(plan.placements.get(4)?.lift).toBe(0)
    expect(plan.placements.get(2)?.lift).toBe(-200)
    expect(plan.breaks).toEqual([])
  })

  it('a quebra de coluna passa o resto para a coluna seguinte', () => {
    const blocks = stack([100, 100, 100]).map((block, index) =>
      index === 0 ? { ...block, columnBreakAfter: true } : block,
    )
    const plan = paginateSections(blocks, [flow(1000, { columns: 2 })])
    expect(columnsOf(plan, 3)).toEqual([0, 1, 1])
    expect(plan.placements.get(1)?.lift).toBe(-100)
  })
})

describe('notas de rodapé (M11)', () => {
  const page = [{ height: 1000, newSheet: false, parity: null, restart: null }] as const
  const SEPARATOR = 20
  /** Uma nota de `lines` linhas de 20 px, com a referência no pé de `at`. */
  const note = (id: string, at: number, lines = 1): MeasuredNote => ({
    id,
    at,
    height: lines * 20,
    lines: Array.from({ length: lines }, (_, line) => line * 20),
  })
  /** Um parágrafo de linhas de 20 px, cortável entre elas. */
  const paragraph = (top: number, lines: number, notes: MeasuredNote[] = []): MeasuredBlock => ({
    top,
    height: lines * 20,
    breakpoints: Array.from({ length: lines - 1 }, (_, line) => top + (line + 1) * 20),
    isPageBreak: false,
    breakAfter: false,
    keepWithNext: false,
    notes,
  })

  it('reserva no pé da folha o separador e a nota', () => {
    // 50 linhas cabem (1000); com a nota de uma linha (20 + 20 de separador)
    // na primeira, só 48.
    const plan = paginateSections([paragraph(0, 60, [note('a', 20)])], page, { separator: SEPARATOR })
    expect(plan.breaks).toEqual([960])
    expect(plan.notes[0]).toEqual([{ id: 'a', fromLine: 0, toLine: 1 }])
    expect(plan.noteHeights[0]).toBe(40)
    expect(plan.notes[1]).toEqual([])
  })

  it('o pedaço da nota vai do topo de uma linha ao da seguinte, e a última até o pé', () => {
    const measured = { height: 70, lines: [0, 22, 40] }
    expect(noteLineTop(measured, 1)).toBe(22)
    expect(noteLineTop(measured, 3)).toBe(70)
    expect(noteSpan(measured, 1, 3)).toBe(48)
    expect(noteSpan(measured, 0, 1)).toBe(22)
  })

  it('sem nota, nada muda', () => {
    const blocks = [paragraph(0, 60)]
    expect(paginateSections(blocks, page, { separator: SEPARATOR }).breaks).toEqual(paginate(blocks, 1000))
  })

  it('a nota que não cabe leva a linha da referência para a folha seguinte', () => {
    // A referência está na linha 49 (pé em 980): o texto caberia, a primeira
    // linha da nota não — a linha desce, e a nota vai com ela.
    const plan = paginateSections([paragraph(0, 60, [note('a', 980)])], page, { separator: SEPARATOR })
    expect(plan.breaks).toEqual([960])
    expect(plan.notes[0]).toEqual([])
    expect(plan.notes[1]).toEqual([{ id: 'a', fromLine: 0, toLine: 1 }])
  })

  it('a nota longa fica com a primeira linha na folha da referência e continua na seguinte', () => {
    const blocks = [paragraph(0, 30, [note('a', 400, 40)]), paragraph(600, 25)]
    const plan = paginateSections(blocks, page, { separator: SEPARATOR })
    // A última nota só precisa da primeira linha: o texto continua até 960, e
    // a nota fica com o que sobra (20 px, uma linha).
    expect(plan.breaks).toEqual([960])
    expect(plan.notes[0]).toEqual([{ id: 'a', fromLine: 0, toLine: 1 }])
    // Na folha seguinte a continuação vem antes de tudo.
    expect(plan.notes[1]).toEqual([{ id: 'a', fromLine: 1, toLine: 40 }])
  })

  it('a continuação longa enche o pé das folhas seguintes, sem tomar a folha toda', () => {
    // A nota de 120 linhas na primeira linha e texto de sobra depois: a
    // continuação pede meia folha em cada uma, e não uma linha por folha.
    const plan = paginateSections([paragraph(0, 1, [note('a', 20, 120)]), paragraph(20, 150)], page, {
      separator: SEPARATOR,
    })
    const slices = plan.notes.map((sheet) => sheet.map((slice) => [slice.fromLine, slice.toLine]))
    expect(slices[0]).toEqual([[0, 1]])
    expect(slices[1]).toEqual([[1, 26]])
    expect(slices[2]).toEqual([[26, 51]])
    // Nenhuma folha passa da altura: texto + separador + notas.
    plan.notes.forEach((_, sheet) => {
      const start = sheet === 0 ? 0 : plan.breaks[sheet - 1]!
      const end = plan.breaks[sheet] ?? 3020
      expect(end - start + (plan.noteHeights[sheet] ?? 0)).toBeLessThanOrEqual(1000)
    })
    expect(plan.notes.flat().at(-1)?.toLine).toBe(120)
  })

  it('a nota maior que a folha continua em folhas só de notas no fim', () => {
    const plan = paginateSections([paragraph(0, 2, [note('a', 20, 120)])], page, { separator: SEPARATOR })
    expect(plan.breaks.length).toBe(2)
    expect(plan.notes.map((slices) => slices.map((slice) => [slice.fromLine, slice.toLine]))).toEqual([
      [[0, 47]],
      [[47, 96]],
      [[96, 120]],
    ])
    expect(plan.sheets).toHaveLength(3)
  })

  it('as notas de várias referências somam, e a última pode ser cortada', () => {
    const blocks = [paragraph(0, 45, [note('a', 20, 3), note('b', 880, 5)])]
    const plan = paginateSections(blocks, page, { separator: SEPARATOR })
    // Até 900 de texto: 900 + 20 + 60 de a + primeira linha de b (20) = 1000
    // cabe; a linha seguinte não. A anterior vai inteira.
    expect(plan.breaks).toEqual([900])
    expect(plan.notes[0]).toEqual([
      { id: 'a', fromLine: 0, toLine: 3 },
      { id: 'b', fromLine: 0, toLine: 1 },
    ])
    expect(plan.notes[1]).toEqual([{ id: 'b', fromLine: 1, toLine: 5 }])
  })

  it('cada seção reserva as notas da própria folha', () => {
    const blocks = [
      { ...paragraph(0, 10, [note('a', 20)]), section: 0 },
      { ...paragraph(200, 10, [note('b', 220)]), section: 1 },
    ]
    const plan = paginateSections(
      blocks,
      [
        { height: 1000, newSheet: false, parity: null, restart: null },
        { height: 1000, newSheet: true, parity: null, restart: null },
      ],
      { separator: SEPARATOR },
    )
    expect(plan.breaks).toEqual([200])
    expect(plan.notes).toEqual([[{ id: 'a', fromLine: 0, toLine: 1 }], [{ id: 'b', fromLine: 0, toLine: 1 }]])
  })

  it('as notas de fim, como blocos depois do texto, cortam entre as linhas delas', () => {
    // Quem as põe no fluxo é a medida (`usePagination`): aqui são blocos comuns.
    const blocks = [paragraph(0, 40), paragraph(820, 20)]
    const plan = paginateSections(blocks, page, { separator: SEPARATOR })
    expect(plan.breaks).toEqual([1000])
  })
})
