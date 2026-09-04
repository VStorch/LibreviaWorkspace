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
})
