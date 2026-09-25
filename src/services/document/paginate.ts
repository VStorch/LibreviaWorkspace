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
  /** Topos de linhas/itens onde a página pode recomeçar; vazio é atômico. */
  readonly breakpoints: readonly number[]
  /** É o nó `pageBreak` — a quebra que a pessoa pediu com Ctrl+Enter. */
  readonly isPageBreak: boolean
  /**
   * A folha termina **depois** deste bloco.
   *
   * É a quebra que o Word gravou dentro do parágrafo. Diferente de
   * `isPageBreak`, que é um bloco só dela.
   */
  readonly breakAfter: boolean
  /** `w:keepNext`: não pode ficar sozinho no pé da página. */
  readonly keepWithNext: boolean
  /**
   * `w:keepLines`: as linhas do parágrafo não se separam. Os pontos de corte
   * continuam medidos, e só valem se o parágrafo sozinho for maior que a folha —
   * aí não há como mantê-lo junto, e o Word também o corta.
   */
  readonly keepLines?: boolean
  /**
   * Os pontos de corte são linhas de parágrafo com controle de viúvas e órfãs
   * (`w:widowControl`, ligado por padrão no Word): nenhuma linha fica sozinha
   * no pé nem no topo da folha. Cortar depois da primeira linha deixaria a
   * órfã, antes da última a viúva — são esses dois cortes que saem. Parágrafo
   * de duas ou três linhas fica sem corte e anda inteiro, como no Word.
   */
  readonly widowControl?: boolean
}

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

  // Sem teto de páginas, e por isso o laço precisa terminar sozinho. Ele
  // termina: em cada volta, ou `index` avança, ou `pageStart` cresce estritamente
  // para um topo de bloco ou um dos seus pontos de corte. Há uma quantidade
  // finita dessas posições; um corte interno nunca permite voltar para trás.
  //
  // Havia um teto de quinhentas páginas, que parecia inofensivo e não era: ao
  // ser alcançado, o laço simplesmente parava, e **todo o resto do documento
  // ficava empilhado na última folha**. Uma altura de página perto de zero —
  // margens absurdas, fonte que não carregou — dava quinhentas folhas em um
  // documento de dez, e o que vinha depois desaparecia de vista. Perder conteúdo
  // de vista é pior do que desenhar folhas demais, e quem protege da altura
  // inválida é a guarda de `pageHeight` logo acima.
  while (index < blocks.length) {
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
      // A quebra que o parágrafo carrega vale depois dele — e não vale se não
      // houver mais nada, senão o documento fecha com uma folha em branco.
      if (block.breakAfter && index < blocks.length) {
        breaks.push(bottom)
        pageStart = bottom
      }

      continue
    }

    const breakpoint = usableBreakpoints(block, pageHeight)
      .filter((at) => at > pageStart && at - pageStart <= pageHeight)
      .at(-1)
    if (breakpoint !== undefined) {
      breaks.push(breakpoint)
      pageStart = breakpoint
      continue
    }

    // Nenhuma linha, item ou linha de tabela cabe: a quebra vai para **antes**
    // do bloco que estouraria.
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
      // Sem corte disponível, o restante fica com a folha só para si.
      // O layout aumenta esse papel para conter o bloco atômico; o próximo
      // bloco continua abrindo uma folha nova, como antes.
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

/** Os cortes internos que o bloco aceita, pelas regras de manter junto. */
function usableBreakpoints(block: MeasuredBlock, pageHeight: number): readonly number[] {
  if (block.keepLines === true && block.height <= pageHeight) return []
  if (block.widowControl !== true) return block.breakpoints
  const guarded = block.breakpoints.slice(1, -1)
  // Maior que a folha e sem corte que respeite a regra: corta assim mesmo,
  // que a alternativa seria uma folha esticada além do papel.
  return guarded.length === 0 && block.height > pageHeight ? block.breakpoints : guarded
}
