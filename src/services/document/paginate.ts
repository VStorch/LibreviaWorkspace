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
  /**
   * Altura das linhas de cabeçalho da tabela (`w:tblHeader`), que se repetem
   * no alto de cada folha em que a tabela continua: a folha nova tem esse
   * tanto a menos para o resto da tabela.
   */
  readonly repeatHeight?: number
  /**
   * Cortes que não são entre linhas de texto e por isso não passam pela regra
   * de viúvas e órfãs: o pé da captura ancorada, onde o LibreOffice deixa a
   * linha vazia do parágrafo descer para a folha seguinte enquanto o quadro
   * fica. Já estão também em `breakpoints`.
   */
  readonly freeBreakpoints?: readonly number[]
  /**
   * Quanto do pé do bloco pode passar da folha: a linha vazia do parágrafo de
   * uma captura ancorada. Medido no LibreOffice, ela entra na margem de baixo
   * em vez de levar o quadro — ou ela mesma — para a folha seguinte.
   */
  readonly hangingBottom?: number
  /**
   * A seção do bloco (M9): o índice dela em `SectionFlow[]`. Ausente é a
   * primeira — o documento de uma seção só.
   */
  readonly section?: number
  /** `w:br w:type="column"`: a coluna termina depois deste bloco. */
  readonly columnBreakAfter?: boolean
}

/**
 * O que a paginação precisa saber de uma seção.
 *
 * As medidas já em pixels de tela e já descontadas as faixas: é a altura útil da
 * folha da seção, a mesma conta de `contentHeightMm`.
 */
export interface SectionFlow {
  readonly height: number
  /**
   * A seção abre folha nova: "próxima página", par, ímpar — e também a contínua
   * cujo papel ou orientação difere da anterior, que o Word trata como próxima
   * página, já que uma folha não muda de tamanho no meio.
   */
  readonly newSheet: boolean
  /** A folha que abre a seção precisa ter número par ou ímpar (`w:type` evenPage/oddPage). */
  readonly parity: 'even' | 'odd' | null
  /** O número que a primeira folha da seção recebe (`w:pgNumType/@w:start`). */
  readonly restart: number | null
  /** Quantas colunas (`w:cols`). Ausente é uma. */
  readonly columns?: number
}

/**
 * Onde um bloco de seção com colunas foi posto (M9).
 *
 * O editor continua sendo uma tira só; a coluna é desenhada **levantando** o
 * primeiro bloco de cada coluna até o topo da região (`lift` negativo) e
 * deslocando para o lado todos os blocos dela. Depois da região, o bloco
 * seguinte desce até o pé da coluna mais alta (`lift` positivo).
 */
export interface ColumnPlacement {
  readonly column: number
  /** Deslocamento vertical a somar ao vão do bloco, em pixels. */
  readonly lift: number
}

/** Uma faixa de colunas numa folha: é nela que a linha separadora é desenhada. */
export interface ColumnRegion {
  /** A folha de conteúdo (índice entre as que têm texto). */
  readonly sheet: number
  /** Topo da região, a contar do topo da coluna de texto da folha. */
  readonly top: number
  readonly height: number
  readonly section: number
  readonly columns: number
}

/** Uma folha do documento, na ordem da pilha. */
export interface SheetPlan {
  /** A seção que abre a folha — é dela o papel, a margem e a faixa. */
  readonly section: number
  /**
   * Folha em branco que o Word insere para a seção par ou ímpar cair na folha
   * certa. Não recebe bloco nenhum; conta na numeração, e já tem o papel da
   * seção que vem depois dela.
   */
  readonly blank: boolean
  /** O número impresso na folha. */
  readonly number: number
  /** É a primeira folha da seção — a da "Primeira página diferente". */
  readonly first: boolean
}

/** Os cortes, e as folhas que eles produzem (as em branco incluídas). */
export interface PagePlan {
  readonly breaks: number[]
  readonly sheets: SheetPlan[]
  /** Os blocos postos em coluna, por índice; os outros não têm entrada. */
  readonly placements: Map<number, ColumnPlacement>
  readonly regions: ColumnRegion[]
}

/**
 * Pontos de corte, em coordenadas de fluxo.
 *
 * Cada valor é onde uma página nova começa. Lista vazia é documento de uma
 * página só.
 */
export function paginate(blocks: readonly MeasuredBlock[], pageHeight: number): number[] {
  return paginateSections(blocks, [{ height: pageHeight, newSheet: false, parity: null, restart: null }])
    .breaks
}

/**
 * Os cortes e as folhas de um documento com seções.
 *
 * A folha tem a altura da seção que a abre: é o papel dela. A seção que começa
 * em folha nova corta antes do primeiro bloco dela — a menos que a folha ainda
 * esteja vazia, e aí a folha passa a ser dela —, e a de página par ou ímpar
 * ganha antes uma folha em branco quando o número não bate, como no Word.
 */
export function paginateSections(
  blocks: readonly MeasuredBlock[],
  sections: readonly SectionFlow[],
): PagePlan {
  const breaks: number[] = []
  const sheets: SheetPlan[] = []
  const flowOf = (section: number): SectionFlow =>
    sections[section] ?? sections.at(-1) ?? { height: 0, newSheet: false, parity: null, restart: null }
  const sectionOf = (block: MeasuredBlock | undefined, fallback: number): number => block?.section ?? fallback

  // A folha nova da seção `section`: numerada a partir da anterior, ou do
  // reinício quando é a primeira da seção. A paridade só vale para a primeira
  // folha de uma seção que a pede, e nunca para a primeira do documento.
  const open = (section: number): void => {
    const previous = sheets.at(-1)
    const first = previous === undefined || previous.section !== section
    const flow = flowOf(section)
    let number = first && flow.restart !== null ? flow.restart : (previous?.number ?? 0) + 1
    if (
      previous !== undefined &&
      first &&
      flow.parity !== null &&
      (number % 2 === 0) !== (flow.parity === 'even')
    ) {
      sheets.push({ section, blank: true, number, first: false })
      number += 1
    }
    sheets.push({ section, blank: false, number, first })
  }

  // A folha que acabou de abrir, vazia, passa a ser da seção que começa nela:
  // refeita, com a numeração e a paridade da seção nova.
  const retarget = (section: number): void => {
    const last = sheets.at(-1)
    if (last === undefined || last.section === section) return
    sheets.pop()
    while (sheets.at(-1)?.blank === true) sheets.pop()
    open(section)
  }

  const placements = new Map<number, ColumnPlacement>()
  const regions: ColumnRegion[] = []
  // O bloco que desceu até o pé da região de colunas: se a folha acabar
  // justamente antes dele, a descida não vale — quem o põe no lugar é o corte.
  let pendingLift: number | null = null

  const firstSection = sectionOf(blocks[0], 0)
  open(firstSection)
  if (blocks.length === 0) return { breaks, sheets, placements, regions }

  let current = firstSection
  let pageHeight = flowOf(firstSection).height
  const cut = (at: number, section: number): void => {
    if (pendingLift !== null && at <= blocks[pendingLift]!.top) placements.delete(pendingLift)
    pendingLift = null
    breaks.push(at)
    open(section)
    pageHeight = flowOf(section).height
  }

  // `pageStart` é de onde a folha conta a altura; `floor`, o último corte.
  // Só diferem quando a folha abre com o cabeçalho repetido de uma tabela: a
  // conta começa acima do corte, pela altura do cabeçalho, mas nada pode voltar
  // para antes do corte.
  let pageStart = 0
  let floor = 0
  let index = 0

  // Sem teto de páginas, e por isso o laço precisa terminar sozinho. Ele
  // termina: em cada volta, ou `index` avança, ou `floor` cresce estritamente
  // para um topo de bloco ou um dos seus pontos de corte. Há uma quantidade
  // finita dessas posições; um corte interno nunca permite voltar para trás.
  //
  // Havia um teto de quinhentas páginas, que parecia inofensivo e não era: ao
  // ser alcançado, o laço simplesmente parava, e **todo o resto do documento
  // ficava empilhado na última folha**. Uma altura de página perto de zero —
  // margens absurdas, fonte que não carregou — dava quinhentas folhas em um
  // documento de dez, e o que vinha depois desaparecia de vista. Perder conteúdo
  // de vista é pior do que desenhar folhas demais, e quem protege da altura
  // inválida é a guarda de `pageHeight` logo abaixo.
  while (index < blocks.length) {
    const block = blocks[index]!

    // A seção nova que começa em folha nova corta antes do primeiro bloco
    // dela. Com a folha ainda vazia — a seção anterior terminou numa quebra de
    // página —, não há o que cortar: a folha passa a ser da seção nova.
    const section = sectionOf(block, current)
    if (section !== current) {
      current = section
      if (flowOf(section).newSheet) {
        if (block.top > floor) {
          cut(block.top, section)
          pageStart = floor = block.top
        } else {
          retarget(section)
          pageHeight = flowOf(section).height
        }
      }
    }

    if (pageHeight <= 0) {
      index += 1
      continue
    }

    // Seção com colunas: os blocos dela, inteiros, vão para as colunas desta
    // folha; o que não couber abre a folha seguinte.
    const columns = flowOf(section).columns ?? 1
    if (columns > 1 && !block.isPageBreak) {
      index = layoutColumns(index, section, columns)
      continue
    }

    // A quebra pedida à mão vale mesmo com a página pela metade, e é por isso
    // que ela vem antes de qualquer conta de altura. O medidor anterior a
    // ignorava, e num documento com capa e sumário a marca caía sempre no lugar
    // errado — três páginas viravam uma.
    if (block.isPageBreak) {
      const after = block.top + block.height
      if (after > floor) {
        cut(after, current)
        pageStart = floor = after
      }

      index += 1
      continue
    }

    const bottom = block.top + block.height
    if (bottom - pageStart <= pageHeight + Math.min(block.hangingBottom ?? 0, pageHeight / 2)) {
      index += 1
      // A quebra que o parágrafo carrega vale depois dele — e não vale se não
      // houver mais nada, senão o documento fecha com uma folha em branco.
      if (block.breakAfter && index < blocks.length) {
        cut(bottom, current)
        pageStart = floor = bottom
      }

      continue
    }

    const breakpoint = usableBreakpoints(block, pageHeight)
      .filter((at) => at > floor && at - pageStart <= pageHeight)
      .at(-1)
    if (breakpoint !== undefined) {
      cut(breakpoint, current)
      floor = breakpoint
      // Cabeçalho maior que meia folha não se repete: repeti-lo deixaria a
      // folha sem lugar para a linha que ele apresenta.
      const repeat = block.repeatHeight ?? 0
      pageStart = repeat > 0 && repeat < pageHeight / 2 ? breakpoint - repeat : breakpoint
      continue
    }

    // Nenhuma linha, item ou linha de tabela cabe: a quebra vai para **antes**
    // do bloco que estouraria.
    let breakAt = block.top
    let opening = current

    // Um título sozinho no pé da página desce junto com o que ele apresenta.
    let candidate = index
    while (candidate > 0) {
      const previous = blocks[candidate - 1]
      if (previous === undefined || !previous.keepWithNext) break
      if (previous.top <= floor) break
      // Não atravessa a quebra de seção que abre folha: o título da seção de
      // cima não desce para a folha da seção de baixo.
      const previousSection = sectionOf(previous, current)
      if (previousSection !== current && flowOf(current).newSheet) break
      candidate -= 1
      breakAt = previous.top
      opening = previousSection
    }

    if (breakAt <= floor) {
      // Sem corte disponível, o restante fica com a folha só para si.
      // O layout aumenta esse papel para conter o bloco atômico; o próximo
      // bloco continua abrindo uma folha nova, como antes.
      pageStart = floor = bottom
      index += 1
      if (index < blocks.length) cut(bottom, current)
      continue
    }

    cut(breakAt, opening)
    pageStart = floor = breakAt
    // `index` não avança: o mesmo bloco é reavaliado na página nova.
  }

  return { breaks, sheets, placements, regions }

  /**
   * Distribui nas colunas desta folha os blocos da seção a partir de `start`, e
   * devolve o primeiro que ficou de fora.
   *
   * Por bloco inteiro: a linha de um parágrafo não sobe para a coluna seguinte,
   * ele vai todo — aproximação do Word, que corta entre linhas. A região termina
   * com a seção, com a folha cheia ou com uma quebra de página. Antes de uma
   * seção contínua na mesma folha as colunas são **equilibradas**, como no Word:
   * a menor altura em que tudo cabe.
   */
  function layoutColumns(start: number, section: number, count: number): number {
    let end = start
    while (end < blocks.length && sectionOf(blocks[end], section) === section) end += 1

    const first = blocks[start]!
    const offset = first.top - pageStart
    const available = pageHeight - offset
    // A região que começa no meio da folha e não comporta nem o primeiro bloco
    // vai para a folha seguinte.
    if (offset > 0 && first.height > available) {
      cut(first.top, section)
      pageStart = floor = first.top
      return start
    }

    let fill = fillColumns(blocks, start, end, available, count)
    const next = blocks[end]
    const balances =
      fill.stop === end && !fill.forced && next !== undefined && !flowOf(sectionOf(next, section)).newSheet
    if (balances) {
      let low = Math.max(...blocks.slice(start, end).map((block) => block.height), 1)
      let high = available
      for (let step = 0; step < 24 && high - low > 0.5; step++) {
        const middle = (low + high) / 2
        const trial = fillColumns(blocks, start, end, middle, count)
        if (trial.stop === end) high = middle
        else low = middle
      }
      fill = fillColumns(blocks, start, end, high, count)
    }

    // O primeiro bloco de cada coluna sobe até o topo da região; os outros a
    // acompanham, porque a tira continua a mesma dentro da coluna.
    let height = 0
    for (const column of fill.columns) {
      const top = blocks[column.from]!
      const last = blocks[column.to - 1]!
      height = Math.max(height, last.top + last.height - top.top)
      const drawn = top.top - pageStart
      const lift = column.index === 0 ? 0 : offset - drawn
      for (let at = column.from; at < column.to; at++) {
        placements.set(at, { column: column.index, lift: at === column.from ? lift : 0 })
      }
      pageStart -= lift
    }
    regions.push({ sheet: breaks.length, top: offset, height, section, columns: count })

    const stop = fill.stop
    const after = blocks[stop]
    if (after === undefined) return stop

    if (fill.forced || stop < end) {
      // Folha cheia, ou quebra de página ou de coluna na última coluna.
      const lastPlaced = blocks[stop - 1]!
      const at = fill.forced ? lastPlaced.top + lastPlaced.height : after.top
      cut(at, sectionOf(after, section))
      pageStart = floor = at
      return stop
    }

    // A seção acabou nesta folha: o bloco seguinte desce ao pé da coluna mais
    // alta, com o espaço natural que ele já tinha acima de si.
    const lastPlaced = blocks[stop - 1]!
    const gap = Math.max(after.top - (lastPlaced.top + lastPlaced.height), 0)
    const lift = offset + height + gap - (after.top - pageStart)
    placements.set(stop, { column: 0, lift })
    pageStart -= lift
    pendingLift = stop
    return stop
  }
}

/** As colunas preenchidas até a altura `height`, bloco inteiro por bloco inteiro. */
function fillColumns(
  blocks: readonly MeasuredBlock[],
  start: number,
  end: number,
  height: number,
  count: number,
): { stop: number; forced: boolean; columns: { index: number; from: number; to: number }[] } {
  const columns: { index: number; from: number; to: number }[] = [{ index: 0, from: start, to: start }]
  const column = (): { index: number; from: number; to: number } => columns.at(-1)!
  const done = (stop: number, forced: boolean) => ({
    stop,
    forced,
    columns: columns.filter((item) => item.to > item.from),
  })
  for (let at = start; at < end; at++) {
    const block = blocks[at]!
    const opened = blocks[column().from]!
    if (at > column().from && block.top + block.height - opened.top > height) {
      if (column().index + 1 >= count) return done(at, false)
      columns.push({ index: column().index + 1, from: at, to: at })
    }
    column().to = at + 1
    if (block.breakAfter === true && at + 1 < blocks.length) return done(at + 1, true)
    if (block.columnBreakAfter === true && at + 1 < end) {
      if (column().index + 1 >= count) return done(at + 1, true)
      columns.push({ index: column().index + 1, from: at + 1, to: at + 1 })
    }
  }
  return done(end, false)
}

/** Os cortes internos que o bloco aceita, pelas regras de manter junto. */
function usableBreakpoints(block: MeasuredBlock, pageHeight: number): readonly number[] {
  if (block.keepLines === true && block.height <= pageHeight) return []
  if (block.widowControl !== true) return block.breakpoints
  const free = block.freeBreakpoints ?? []
  const lines = block.breakpoints.filter((at) => !free.includes(at))
  const guarded = [...lines.slice(1, -1), ...free].sort((left, right) => left - right)
  // Maior que a folha e sem corte que respeite a regra: corta assim mesmo,
  // que a alternativa seria uma folha esticada além do papel.
  return guarded.length === 0 && block.height > pageHeight ? block.breakpoints : guarded
}
