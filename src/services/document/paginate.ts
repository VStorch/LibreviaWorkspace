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

  const notes: NoteSlice[][] = []
  const noteHeights: number[] = []

  const firstSection = sectionOf(blocks[0], 0)
  open(firstSection)
  if (blocks.length === 0) return { breaks, sheets, placements, regions, notes, noteHeights }

  // A folha leva a nota cuja referência ela leva. A nota longa segue o Word: a
  // linha da referência e pelo menos a primeira linha da nota ficam na mesma
  // folha, e o resto continua no alto da área de notas seguinte (`carry`).
  const footnotes = blocks.flatMap((block) => block.notes ?? [])
  let nextNote = 0
  let carry: { note: MeasuredNote; from: number }[] = []
  const separator = noteFlow.separator

  // As notas que a folha levaria se terminasse em `at`: as que continuam da
  // anterior e as das referências até ali.
  const pendingNotes = (at: number): { note: MeasuredNote; from: number }[] => {
    const list = [...carry]
    for (let next = nextNote; next < footnotes.length && footnotes[next]!.at <= at + 0.5; next++) {
      list.push({ note: footnotes[next]!, from: 0 })
    }
    return list
  }

  // O espaço que as notas pedem para a folha terminar em `at`: as novas
  // inteiras, menos a última, de que basta a primeira linha. Cresce com `at`, e
  // por isso "o último corte que cabe" continua valendo. A continuação vem antes
  // do texto, como no Word, e pede o resto inteiro até meia folha.
  const noteNeed = (at: number): number => {
    if (carry.length === 0 && (nextNote >= footnotes.length || footnotes[nextNote]!.at > at + 0.5)) return 0
    const fresh = pendingNotes(at).slice(carry.length)
    let carried = 0
    for (const item of carry) carried += noteSpan(item.note, item.from, lineCount(item.note))
    const first = carry[0]
    let need =
      separator +
      (first === undefined
        ? 0
        : Math.max(Math.min(carried, pageHeight / 2), noteSpan(first.note, first.from, first.from + 1)))
    fresh.forEach((item, position) => {
      need +=
        position === fresh.length - 1
          ? noteSpan(item.note, item.from, item.from + 1)
          : noteSpan(item.note, item.from, lineCount(item.note))
    })
    return need
  }

  // Fecha a folha que termina em `at` com `used` de texto: as notas que cabem
  // vão inteiras, a primeira que não cabe é cortada entre linhas, e o resto
  // continua na folha seguinte.
  const settleNotes = (at: number, used: number): void => {
    const list = pendingNotes(at)
    while (nextNote < footnotes.length && footnotes[nextNote]!.at <= at + 0.5) nextNote += 1
    carry = []
    const placed: NoteSlice[] = []
    let room = pageHeight - used - separator
    let height = 0
    for (const item of list) {
      if (carry.length > 0) {
        carry.push(item)
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
      if (to < total) carry.push({ note: item.note, from: to })
    }
    notes[breaks.length] = placed
    noteHeights[breaks.length] = placed.length > 0 ? height + separator : 0
  }

  let current = firstSection
  let pageHeight = flowOf(firstSection).height
  const cut = (at: number, section: number, used = at - pageStart): void => {
    if (pendingLift !== null && at <= blocks[pendingLift]!.top) placements.delete(pendingLift)
    pendingLift = null
    settleNotes(at, used)
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

  // Sem teto de páginas: em cada volta `index` avança ou `floor` cresce
  // estritamente, e há uma quantidade finita dessas posições. Um teto pararia o
  // laço e empilharia o resto do documento na última folha; quem protege da
  // altura inválida é a guarda de `pageHeight`.
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

    // A quebra pedida à mão vale mesmo com a página pela metade, por isso vem
    // antes de qualquer conta de altura.
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
    if (
      bottom - pageStart + noteNeed(bottom) <=
      pageHeight + Math.min(block.hangingBottom ?? 0, pageHeight / 2)
    ) {
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
      .filter((at) => at > floor && at - pageStart + noteNeed(at) <= pageHeight)
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
      // Sem corte disponível, o bloco atômico fica com a folha só para si, e o
      // layout aumenta o papel para contê-lo.
      const used = bottom - pageStart
      pageStart = floor = bottom
      index += 1
      if (index < blocks.length) cut(bottom, current, used)
      continue
    }

    cut(breakAt, opening)
    pageStart = floor = breakAt
    // `index` não avança: o mesmo bloco é reavaliado na página nova.
  }

  // A última folha fecha com as notas que sobraram; a nota que ainda não coube
  // continua em folhas só de notas, depois do texto.
  const end = blocks.reduce((bottom, block) => Math.max(bottom, block.top + block.height), 0)
  let used = Math.max(end - pageStart, 0)
  for (;;) {
    settleNotes(Number.POSITIVE_INFINITY, used)
    if (carry.length === 0 || pageHeight <= 0) break
    breaks.push(end)
    open(current)
    pageStart = floor = end
    used = 0
  }

  return { breaks, sheets, placements, regions, notes, noteHeights }

  /**
   * Distribui nas colunas desta folha os blocos da seção a partir de `start`, e
   * devolve o primeiro que ficou de fora. Por bloco inteiro, aproximando o Word,
   * que corta entre linhas. Antes de uma seção contínua na mesma folha as colunas
   * são equilibradas, como no Word.
   */
  function layoutColumns(start: number, section: number, count: number): number {
    let end = start
    while (end < blocks.length && sectionOf(blocks[end], section) === section) end += 1

    const first = blocks[start]!
    const offset = first.top - pageStart
    // As notas da região saem da altura das colunas; a área delas fica embaixo,
    // na largura da folha (limitação declarada: o Word as põe sob cada coluna).
    const last = blocks[end - 1]!
    const available = pageHeight - offset - noteNeed(last.top + last.height)
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
