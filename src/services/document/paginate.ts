/**
 * Onde a página termina.
 *
 * Recebe blocos **já medidos** e devolve os pontos de corte. Não toca no DOM:
 * quem mede é o editor, que é o único que sabe a altura real de cada bloco
 * depois da fonte carregar. Separar as duas coisas é o que torna a regra
 * testável — a medição precisa de um navegador, a decisão não.
 *
 * As posições são em **coordenadas de fluxo**: a altura que o bloco teria se o
 * documento fosse uma tira contínua, sem os vãos entre as folhas. É o que
 * resolve a realimentação — inserir os vãos muda o `offsetTop` de tudo que vem
 * depois, e recalcular sobre a nova medida entraria em laço. Em coordenadas de
 * fluxo o cálculo é um passo só, e quem desenha soma os vãos depois.
 */

/** Um bloco de primeiro nível, medido na tela. */
export interface MeasuredBlock {
  /** Topo em coordenadas de fluxo. */
  readonly top: number
  readonly height: number
  /** É o nó `pageBreak` — a quebra que a pessoa pediu com Ctrl+Enter. */
  readonly isPageBreak: boolean
  /** `w:keepNext`: não pode ficar sozinho no pé da página. */
  readonly keepWithNext: boolean
}

/**
 * Teto de páginas.
 *
 * Um documento não tem quinhentas páginas por acidente, mas uma altura de
 * página perto de zero — margens absurdas, fonte que não carregou — teria. O
 * teto troca um travamento por um documento estranho, que é o erro mais barato.
 */
const MAX_PAGES = 500

/**
 * Pontos de corte, em coordenadas de fluxo.
 *
 * Cada valor é onde uma página nova começa. Lista vazia é documento de uma
 * página só.
 */
export function paginate(blocks: readonly MeasuredBlock[], pageHeight: number): number[] {
  if (pageHeight <= 0) return []

  const breaks: number[] = []
  let pageStart = 0
  let index = 0

  while (index < blocks.length && breaks.length < MAX_PAGES) {
    const block = blocks[index]!

    // A quebra pedida à mão vale mesmo com a página pela metade, e é por isso
    // que ela vem antes de qualquer conta de altura. O medidor anterior a
    // ignorava, e num documento com capa e sumário a marca caía sempre no lugar
    // errado — três páginas viravam uma.
    if (block.isPageBreak) {
      const after = block.top + block.height
      if (after > pageStart) {
        breaks.push(after)
        pageStart = after
      }

      index += 1
      continue
    }

    const bottom = block.top + block.height
    if (bottom - pageStart <= pageHeight) {
      index += 1
      continue
    }

    // Quebra **antes** do bloco que estouraria — a mesma decisão do navegador ao
    // imprimir, e o motivo de a página nunca cortar um parágrafo ao meio.
    let breakAt = block.top

    // Um título sozinho no pé da página desce junto com o que ele apresenta.
    let candidate = index
    while (candidate > 0) {
      const previous = blocks[candidate - 1]
      if (previous === undefined || !previous.keepWithNext) break
      if (previous.top <= pageStart) break
      candidate -= 1
      breakAt = previous.top
    }

    if (breakAt <= pageStart) {
      // Bloco mais alto que uma página inteira — uma captura de tela grande.
      // Não há onde parti-lo, então ele fica com a folha só para si e
      // transborda. O que vem depois começa numa folha nova: deixar o próximo
      // parágrafo encostado embaixo do transbordo o poria fora do papel
      // desenhado, que parece defeito de desenho e não documento grande demais.
      pageStart = bottom
      index += 1
      if (index < blocks.length) breaks.push(bottom)
      continue
    }

    breaks.push(breakAt)
    pageStart = breakAt
    // `index` não avança: o mesmo bloco é reavaliado na página nova.
  }

  return breaks
}
