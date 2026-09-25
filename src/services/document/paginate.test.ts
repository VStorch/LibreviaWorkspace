import { describe, expect, it } from 'vitest'
import { paginate, type MeasuredBlock } from './paginate.js'

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
      // Antes daqui o parágrafo descia inteiro e deixava 300 de buraco na folha.
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
  })
})
