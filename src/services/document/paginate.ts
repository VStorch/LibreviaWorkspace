/**
 * Recebe blocos **já medidos** pelo editor e devolve os pontos de corte, sem
 * tocar no DOM.
 *
 * As posições são em **coordenadas de fluxo**: a altura que o bloco teria numa
 * tira contínua, sem os vãos entre as folhas. Inserir os vãos muda o
 * `offsetTop` de tudo que vem depois; em coordenadas de fluxo o cálculo é um
 * passo só, e quem desenha soma os vãos depois.
 */

export interface MeasuredBlock {
  /** Topo em coordenadas de fluxo. */
  readonly top: number
  readonly height: number
  /** Topos de linhas/itens onde a página pode recomeçar; vazio é atômico. */
  readonly breakpoints: readonly number[]
  /** É o nó `pageBreak` — a quebra que a pessoa pediu com Ctrl+Enter. */
  readonly isPageBreak: boolean
  /** A quebra que o Word gravou dentro do parágrafo: a folha termina depois do bloco. */
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
   * `w:widowControl`, ligado por padrão no Word: os cortes depois da primeira
   * linha e antes da última saem, e parágrafo de até três linhas anda inteiro.
   */
  readonly widowControl?: boolean
  /**
   * Altura das linhas de cabeçalho da tabela (`w:tblHeader`), que se repetem
   * no alto de cada folha em que a tabela continua: a folha nova tem esse
   * tanto a menos para o resto da tabela.
   */
  readonly repeatHeight?: number
  /**
   * Cortes fora da regra de viúvas e órfãs: o pé da captura ancorada, onde o
   * LibreOffice deixa a linha vazia do parágrafo descer enquanto o quadro fica.
   * Já estão também em `breakpoints`.
   */
  readonly freeBreakpoints?: readonly number[]
  /** A linha vazia do parágrafo de uma captura ancorada, que o LibreOffice deixa entrar na margem de baixo. */
  readonly hangingBottom?: number
  /**
   * A seção do bloco: o índice dela em `SectionFlow[]`. Ausente é a primeira —
   * o documento de uma seção só.
   */
  readonly section?: number
  /** `w:br w:type="column"`: a coluna termina depois deste bloco. */
  readonly columnBreakAfter?: boolean
  /** As notas de rodapé cujas referências estão neste bloco, na ordem do texto. */
  readonly notes?: readonly MeasuredNote[]
}

/** A altura não depende da paginação: o corpo é medido fora do fluxo, na largura da coluna. */
export interface MeasuredNote {
  readonly id: string
  /** O pé da linha da referência, em coordenadas de fluxo. */
  readonly at: number
  readonly height: number
  /** O topo de cada linha do corpo, a partir do topo dele; a primeira é 0. */
  readonly lines: readonly number[]
}

/** O pedaço de uma nota que cai numa folha: as linhas `[fromLine, toLine)`. */
export interface NoteSlice {
  readonly id: string
  readonly fromLine: number
  readonly toLine: number
}

/** Medidas em pixels de tela, já sem as faixas: a altura útil de `contentHeightMm`. */
export interface SectionFlow {
  readonly height: number
  /** Também a contínua com outro papel ou orientação, que o Word trata como próxima página. */
  readonly newSheet: boolean
  /** A folha que abre a seção precisa ter número par ou ímpar (`w:type` evenPage/oddPage). */
  readonly parity: 'even' | 'odd' | null
  /** O número que a primeira folha da seção recebe (`w:pgNumType/@w:start`). */
  readonly restart: number | null
  /** Quantas colunas (`w:cols`). Ausente é uma. */
  readonly columns?: number
}

/**
 * O editor continua sendo uma tira só: a coluna é desenhada **levantando** o
 * primeiro bloco de cada coluna até o topo da região (`lift` negativo) e
 * deslocando os blocos dela para o lado. Depois da região, o bloco seguinte
 * desce até o pé da coluna mais alta (`lift` positivo).
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

export interface SheetPlan {
  /** A seção que abre a folha — é dela o papel, a margem e a faixa. */
  readonly section: number
  /**
   * Folha em branco que o Word insere para a seção par ou ímpar cair na folha
   * certa. Conta na numeração e tem o papel da seção seguinte.
   */
  readonly blank: boolean
  readonly number: number
  /** A folha da "Primeira página diferente". */
  readonly first: boolean
}

/** Os cortes, e as folhas que eles produzem (as em branco incluídas). */
export interface PagePlan {
  readonly breaks: number[]
  readonly sheets: SheetPlan[]
  /** Os blocos postos em coluna, por índice; os outros não têm entrada. */
  readonly placements: Map<number, ColumnPlacement>
  readonly regions: ColumnRegion[]
  /** As notas de rodapé de cada folha de conteúdo (índice entre as com texto). */
  readonly notes: NoteSlice[][]
  /** A altura da área de notas de cada folha de conteúdo, com o separador; 0 sem nota. */
  readonly noteHeights: number[]
}

/** O que a paginação precisa saber das notas, além dos blocos. */
export interface NoteFlow {
  /** A altura do separador entre o texto e as notas. */
  readonly separator: number
}

/** O topo da linha `line` da nota; a linha depois da última é o pé dela. */
export function noteLineTop(note: Pick<MeasuredNote, 'height' | 'lines'>, line: number): number {
  return line >= note.lines.length ? note.height : (note.lines[line] ?? 0)
}

/** A altura das linhas `[from, to)` da nota. */
export function noteSpan(note: Pick<MeasuredNote, 'height' | 'lines'>, from: number, to: number): number {
  return noteLineTop(note, to) - noteLineTop(note, from)
}

function lineCount(note: MeasuredNote): number {
  return Math.max(note.lines.length, 1)
}

/** Cada valor é onde uma página nova começa; lista vazia é documento de uma página. */
export function paginate(blocks: readonly MeasuredBlock[], pageHeight: number): number[] {
  return paginateSections(blocks, [{ height: pageHeight, newSheet: false, parity: null, restart: null }])
    .breaks
}

/**
 * A folha tem a altura da seção que a abre. A seção que começa em folha nova
 * corta antes do primeiro bloco dela, a menos que a folha esteja vazia; a de
 * página par ou ímpar ganha antes uma folha em branco quando o número não bate,
 * como no Word.
 */
export function paginateSections(
  blocks: readonly MeasuredBlock[],
  sections: readonly SectionFlow[],
  noteFlow: NoteFlow = { separator: 0 },
): PagePlan {
  return new Paginator(blocks, sections, noteFlow.separator).run()
}

interface PendingNote {
  readonly note: MeasuredNote
  readonly from: number
}

type ColumnFill = ReturnType<typeof fillColumns>

const sectionOf = (block: MeasuredBlock | undefined, fallback: number): number => block?.section ?? fallback

class Paginator {
  private readonly breaks: number[] = []
  private readonly sheets: SheetPlan[] = []
  private readonly placements = new Map<number, ColumnPlacement>()
  private readonly regions: ColumnRegion[] = []
  private readonly notes: NoteSlice[][] = []
  private readonly noteHeights: number[] = []

  /**
   * O bloco que desceu até o pé da região de colunas: se a folha acabar
   * justamente antes dele, a descida não vale — quem o põe no lugar é o corte.
   */
  private pendingLift: number | null = null

  /**
   * A folha leva a nota cuja referência ela leva. A nota longa segue o Word: a
   * linha da referência e pelo menos a primeira linha da nota ficam na mesma
   * folha, e o resto continua no alto da área de notas seguinte (`carry`).
   */
  private readonly footnotes: readonly MeasuredNote[]
  private nextNote = 0
  private carry: PendingNote[] = []

  private current: number
  private pageHeight: number

  /**
   * `pageStart` é de onde a folha conta a altura; `floor`, o último corte.
   * Só diferem quando a folha abre com o cabeçalho repetido de uma tabela: a
   * conta começa acima do corte, pela altura do cabeçalho, mas nada pode voltar
   * para antes do corte.
   */
  private pageStart = 0
  private floor = 0

  constructor(
    private readonly blocks: readonly MeasuredBlock[],
    private readonly sections: readonly SectionFlow[],
    private readonly separator: number,
  ) {
    this.footnotes = blocks.flatMap((block) => block.notes ?? [])
    this.current = sectionOf(blocks[0], 0)
    this.pageHeight = this.flowOf(this.current).height
  }

  /**
   * Sem teto de páginas: em cada volta o índice avança ou `floor` cresce
   * estritamente, e há uma quantidade finita dessas posições. Um teto pararia o
   * laço e empilharia o resto do documento na última folha; quem protege da
   * altura inválida é a guarda de `pageHeight`.
   */
  run(): PagePlan {
    this.open(this.current)
    if (this.blocks.length === 0) return this.plan()

    let index = 0
    while (index < this.blocks.length) index = this.step(index)
    this.closeNotes()
    return this.plan()
  }

  private plan(): PagePlan {
    const { breaks, sheets, placements, regions, notes, noteHeights } = this
    return { breaks, sheets, placements, regions, notes, noteHeights }
  }

  private flowOf(section: number): SectionFlow {
    return (
      this.sections[section] ??
      this.sections.at(-1) ?? { height: 0, newSheet: false, parity: null, restart: null }
    )
  }

  /** Devolve o próximo bloco a avaliar, que é o mesmo quando a folha virou antes dele. */
  private step(index: number): number {
    const block = this.blocks[index]!
    const section = this.enterSection(block)

    if (this.pageHeight <= 0) return index + 1

    // Seção com colunas: os blocos dela, inteiros, vão para as colunas desta
    // folha; o que não couber abre a folha seguinte.
    const columns = this.flowOf(section).columns ?? 1
    if (columns > 1 && !block.isPageBreak) return this.layoutColumns(index, section, columns)

    // A quebra pedida à mão vale mesmo com a página pela metade, por isso vem
    // antes de qualquer conta de altura.
    if (block.isPageBreak) {
      const after = block.top + block.height
      if (after > this.floor) this.cutAndRestart(after, this.current)
      return index + 1
    }

    const bottom = block.top + block.height
    if (
      bottom - this.pageStart + this.noteNeed(bottom) <=
      this.pageHeight + Math.min(block.hangingBottom ?? 0, this.pageHeight / 2)
    ) {
      // A quebra que o parágrafo carrega vale depois dele — e não vale se não
      // houver mais nada, senão o documento fecha com uma folha em branco.
      if (block.breakAfter && index + 1 < this.blocks.length) this.cutAndRestart(bottom, this.current)
      return index + 1
    }

    if (this.breakInside(block)) return index
    return this.breakBefore(index, block)
  }

  /**
   * A seção nova que começa em folha nova corta antes do primeiro bloco dela.
   * Com a folha ainda vazia — a seção anterior terminou numa quebra de página —,
   * não há o que cortar: a folha passa a ser da seção nova.
   */
  private enterSection(block: MeasuredBlock): number {
    const section = sectionOf(block, this.current)
    if (section === this.current) return section
    this.current = section
    if (!this.flowOf(section).newSheet) return section
    if (block.top > this.floor) {
      this.cutAndRestart(block.top, section)
    } else {
      this.retarget(section)
      this.pageHeight = this.flowOf(section).height
    }
    return section
  }

  private breakInside(block: MeasuredBlock): boolean {
    const breakpoint = usableBreakpoints(block, this.pageHeight)
      .filter((at) => at > this.floor && at - this.pageStart + this.noteNeed(at) <= this.pageHeight)
      .at(-1)
    if (breakpoint === undefined) return false
    this.cut(breakpoint, this.current)
    this.floor = breakpoint
    // Cabeçalho maior que meia folha não se repete: repeti-lo deixaria a
    // folha sem lugar para a linha que ele apresenta.
    const repeat = block.repeatHeight ?? 0
    this.pageStart = repeat > 0 && repeat < this.pageHeight / 2 ? breakpoint - repeat : breakpoint
    return true
  }

  /**
   * Nenhuma linha, item ou linha de tabela cabe: a quebra vai para **antes** do
   * bloco que estouraria.
   */
  private breakBefore(index: number, block: MeasuredBlock): number {
    const { at, opening } = this.keptTogetherStart(index, block)
    if (at <= this.floor) {
      // Sem corte disponível, o bloco atômico fica com a folha só para si, e o
      // layout aumenta o papel para contê-lo.
      const bottom = block.top + block.height
      const used = bottom - this.pageStart
      this.pageStart = this.floor = bottom
      if (index + 1 < this.blocks.length) this.cut(bottom, this.current, used)
      return index + 1
    }
    this.cutAndRestart(at, opening)
    return index
  }

  /** Um título sozinho no pé da página desce junto com o que ele apresenta. */
  private keptTogetherStart(index: number, block: MeasuredBlock): { at: number; opening: number } {
    let at = block.top
    let opening = this.current
    for (let candidate = index; candidate > 0; candidate--) {
      const previous = this.blocks[candidate - 1]
      if (previous === undefined || !previous.keepWithNext) break
      if (previous.top <= this.floor) break
      // Não atravessa a quebra de seção que abre folha: o título da seção de
      // cima não desce para a folha da seção de baixo.
      const previousSection = sectionOf(previous, this.current)
      if (previousSection !== this.current && this.flowOf(this.current).newSheet) break
      at = previous.top
      opening = previousSection
    }
    return { at, opening }
  }

  /**
   * A última folha fecha com as notas que sobraram; a nota que ainda não coube
   * continua em folhas só de notas, depois do texto.
   */
  private closeNotes(): void {
    const end = this.blocks.reduce((bottom, block) => Math.max(bottom, block.top + block.height), 0)
    let used = Math.max(end - this.pageStart, 0)
    for (;;) {
      this.settleNotes(Number.POSITIVE_INFINITY, used)
      if (this.carry.length === 0 || this.pageHeight <= 0) break
      this.breaks.push(end)
      this.open(this.current)
      this.pageStart = this.floor = end
      used = 0
    }
  }

  /**
   * A folha nova da seção `section`: numerada a partir da anterior, ou do
   * reinício quando é a primeira da seção. A paridade só vale para a primeira
   * folha de uma seção que a pede, e nunca para a primeira do documento.
   */
  private open(section: number): void {
    const previous = this.sheets.at(-1)
    const first = previous === undefined || previous.section !== section
    const flow = this.flowOf(section)
    let number = first && flow.restart !== null ? flow.restart : (previous?.number ?? 0) + 1
    if (
      previous !== undefined &&
      first &&
      flow.parity !== null &&
      (number % 2 === 0) !== (flow.parity === 'even')
    ) {
      this.sheets.push({ section, blank: true, number, first: false })
      number += 1
    }
    this.sheets.push({ section, blank: false, number, first })
  }

  /**
   * A folha que acabou de abrir, vazia, passa a ser da seção que começa nela:
   * refeita, com a numeração e a paridade da seção nova.
   */
  private retarget(section: number): void {
    const last = this.sheets.at(-1)
    if (last === undefined || last.section === section) return
    this.sheets.pop()
    while (this.sheets.at(-1)?.blank === true) this.sheets.pop()
    this.open(section)
  }

  private cut(at: number, section: number, used = at - this.pageStart): void {
    if (this.pendingLift !== null && at <= this.blocks[this.pendingLift]!.top)
      this.placements.delete(this.pendingLift)
    this.pendingLift = null
    this.settleNotes(at, used)
    this.breaks.push(at)
    this.open(section)
    this.pageHeight = this.flowOf(section).height
  }

  private cutAndRestart(at: number, section: number): void {
    this.cut(at, section)
    this.pageStart = this.floor = at
  }

  /**
   * As notas que a folha levaria se terminasse em `at`: as que continuam da
   * anterior e as das referências até ali.
   */
  private pendingNotes(at: number): PendingNote[] {
    const list = [...this.carry]
    for (
      let next = this.nextNote;
      next < this.footnotes.length && this.footnotes[next]!.at <= at + 0.5;
      next++
    ) {
      list.push({ note: this.footnotes[next]!, from: 0 })
    }
    return list
  }

  /**
   * O espaço que as notas pedem para a folha terminar em `at`: as novas
   * inteiras, menos a última, de que basta a primeira linha. Cresce com `at`, e
   * por isso "o último corte que cabe" continua valendo. A continuação vem antes
   * do texto, como no Word, e pede o resto inteiro até meia folha.
   */
  private noteNeed(at: number): number {
    const { carry, footnotes, nextNote } = this
    if (carry.length === 0 && (nextNote >= footnotes.length || footnotes[nextNote]!.at > at + 0.5)) return 0
    const fresh = this.pendingNotes(at).slice(carry.length)
    let carried = 0
    for (const item of carry) carried += noteSpan(item.note, item.from, lineCount(item.note))
    const first = carry[0]
    let need =
      this.separator +
      (first === undefined
        ? 0
        : Math.max(Math.min(carried, this.pageHeight / 2), noteSpan(first.note, first.from, first.from + 1)))
    fresh.forEach((item, position) => {
      need +=
        position === fresh.length - 1
          ? noteSpan(item.note, item.from, item.from + 1)
          : noteSpan(item.note, item.from, lineCount(item.note))
    })
    return need
  }

  /**
   * Fecha a folha que termina em `at` com `used` de texto: as notas que cabem
   * vão inteiras, a primeira que não cabe é cortada entre linhas, e o resto
   * continua na folha seguinte.
   */
  private settleNotes(at: number, used: number): void {
    const list = this.pendingNotes(at)
    while (this.nextNote < this.footnotes.length && this.footnotes[this.nextNote]!.at <= at + 0.5)
      this.nextNote += 1
    this.carry = []
    const placed: NoteSlice[] = []
    let room = this.pageHeight - used - this.separator
    let height = 0
    for (const item of list) {
      if (this.carry.length > 0) {
        this.carry.push(item)
        continue
      }
      const total = lineCount(item.note)
      let to = item.from
      while (to < total && noteSpan(item.note, item.from, to + 1) <= room + 0.5) to += 1
      // Pelo menos uma linha na folha que não levou nenhuma: é o que faz a nota
      // maior que a folha terminar, uma folha por vez.
      if (to === item.from && placed.length === 0) to += 1
      if (to > item.from) {
        const span = noteSpan(item.note, item.from, to)
        placed.push({ id: item.note.id, fromLine: item.from, toLine: to })
        room -= span
        height += span
      }
      if (to < total) this.carry.push({ note: item.note, from: to })
    }
    this.notes[this.breaks.length] = placed
    this.noteHeights[this.breaks.length] = placed.length > 0 ? height + this.separator : 0
  }

  /**
   * Distribui nas colunas desta folha os blocos da seção a partir de `start`, e
   * devolve o primeiro que ficou de fora. Por bloco inteiro, aproximando o Word,
   * que corta entre linhas.
   */
  private layoutColumns(start: number, section: number, count: number): number {
    const { blocks } = this
    let end = start
    while (end < blocks.length && sectionOf(blocks[end], section) === section) end += 1

    const first = blocks[start]!
    const offset = first.top - this.pageStart
    // As notas da região saem da altura das colunas; a área delas fica embaixo,
    // na largura da folha (limitação declarada: o Word as põe sob cada coluna).
    const last = blocks[end - 1]!
    const available = this.pageHeight - offset - this.noteNeed(last.top + last.height)
    // A região que começa no meio da folha e não comporta nem o primeiro bloco
    // vai para a folha seguinte.
    if (offset > 0 && first.height > available) {
      this.cutAndRestart(first.top, section)
      return start
    }

    const fill = this.columnFill(start, end, available, count, section)
    const height = this.placeColumns(fill, offset)
    this.regions.push({ sheet: this.breaks.length, top: offset, height, section, columns: count })
    return this.leaveColumns(fill, end, section, offset + height)
  }

  /** Antes de uma seção contínua na mesma folha as colunas são equilibradas, como no Word. */
  private columnFill(
    start: number,
    end: number,
    available: number,
    count: number,
    section: number,
  ): ColumnFill {
    const { blocks } = this
    const fill = fillColumns(blocks, start, end, available, count)
    const next = blocks[end]
    const balances =
      fill.stop === end &&
      !fill.forced &&
      next !== undefined &&
      !this.flowOf(sectionOf(next, section)).newSheet
    if (!balances) return fill

    let low = Math.max(...blocks.slice(start, end).map((block) => block.height), 1)
    let high = available
    for (let step = 0; step < 24 && high - low > 0.5; step++) {
      const middle = (low + high) / 2
      const trial = fillColumns(blocks, start, end, middle, count)
      if (trial.stop === end) high = middle
      else low = middle
    }
    return fillColumns(blocks, start, end, high, count)
  }

  /**
   * O primeiro bloco de cada coluna sobe até o topo da região; os outros a
   * acompanham, porque a tira continua a mesma dentro da coluna. Devolve a
   * altura da coluna mais alta.
   */
  private placeColumns(fill: ColumnFill, offset: number): number {
    let height = 0
    for (const column of fill.columns) {
      const top = this.blocks[column.from]!
      const last = this.blocks[column.to - 1]!
      height = Math.max(height, last.top + last.height - top.top)
      const drawn = top.top - this.pageStart
      const lift = column.index === 0 ? 0 : offset - drawn
      for (let at = column.from; at < column.to; at++) {
        this.placements.set(at, { column: column.index, lift: at === column.from ? lift : 0 })
      }
      this.pageStart -= lift
    }
    return height
  }

  /** @param regionBottom o pé da região de colunas, a contar do topo da folha. */
  private leaveColumns(fill: ColumnFill, end: number, section: number, regionBottom: number): number {
    const stop = fill.stop
    const after = this.blocks[stop]
    if (after === undefined) return stop

    const lastPlaced = this.blocks[stop - 1]!
    if (fill.forced || stop < end) {
      // Folha cheia, ou quebra de página ou de coluna na última coluna.
      const at = fill.forced ? lastPlaced.top + lastPlaced.height : after.top
      this.cutAndRestart(at, sectionOf(after, section))
      return stop
    }

    // A seção acabou nesta folha: o bloco seguinte desce ao pé da coluna mais
    // alta, com o espaço natural que ele já tinha acima de si.
    const gap = Math.max(after.top - (lastPlaced.top + lastPlaced.height), 0)
    const lift = regionBottom + gap - (after.top - this.pageStart)
    this.placements.set(stop, { column: 0, lift })
    this.pageStart -= lift
    this.pendingLift = stop
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
