import { Extension } from '@tiptap/core'
import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Mapping, ReplaceStep, StepMap, canJoin, type Step } from '@tiptap/pm/transform'
import { DELETION, INSERTION, ZERO_WIDTH, blockRevisionOf } from './track-changes.js'

/**
 * Controle de alterações (M10, fase 2): o que se digita com o controle ligado.
 *
 * A transação é reescrita **antes** de ser aplicada (`dispatchTransaction`): sai
 * uma transação só, com a seleção certa e um passo de desfazer. A regra, para
 * cada `ReplaceStep`:
 *
 * - o que entra ganha a marca `insertion` do autor, e perde qualquer `deletion`;
 * - o que sai **fica**, com a marca `deletion` — menos o que já era inserção do
 *   mesmo autor (esse sai de verdade, como no Word) e o que já era exclusão (fica
 *   como estava; o cursor só passa por cima);
 * - o Enter que sai vira a marca de parágrafo excluída (`markRevision del`) do
 *   bloco de cima, e o que entra, a inserida (`markRevision ins`); a marca que o
 *   próprio autor inseriu sai de verdade — os blocos se juntam;
 * - a linha de tabela inteira que sai fica, com `rowRevision del`; a que entra
 *   ganha `rowRevision ins`.
 *
 * Formatação (marcas, atributos, `ReplaceAroundStep` de lista e de tipo de bloco)
 * passa sem controle, como passos comuns. Também passam o desfazer, o aceitar e
 * rejeitar (`SKIP_TRACKING`), o que não entra no histórico e a ponta de comentário
 * ou de marcador sozinha — ponta não é texto.
 *
 * Durante a composição do IME só a marca de inserção entra: devolver o que a
 * composição apaga desmontaria o DOM que o IME está editando. O que ela apaga no
 * meio é, quase sempre, o próprio texto ainda em composição.
 */

/** A meta que tira a transação do controle — aceitar e rejeitar, por exemplo. */
export const SKIP_TRACKING = 'trackChanges:skip'

/** A meta do histórico (`prosemirror-history`): desfazer e refazer não são edição nova. */
const HISTORY_META = 'history$'

/**
 * A data da revisão, com precisão de minuto, como o Word a grava. É ela que deixa
 * o que se digita no mesmo minuto virar um trecho só: as marcas ficam iguais e o
 * texto se funde.
 */
export function revisionDate(now: Date): string {
  return `${now.toISOString().slice(0, 16)}:00Z`
}

/** O autor e a data de uma edição controlada. */
interface Revision {
  readonly author: string
  readonly date: string
}

export interface TrackOptions {
  /** A composição do IME está em curso (`view.composing`). */
  readonly composing?: boolean
}

/** A transação deve passar pelo controle? */
export function shouldTrack(tr: Transaction): boolean {
  if (!tr.docChanged) return false
  if (tr.getMeta(SKIP_TRACKING) === true) return false
  if (tr.getMeta(HISTORY_META) !== undefined) return false
  return tr.getMeta('addToHistory') !== false
}

// --- o que ficou no documento controlado e não está no original --------------

/**
 * O texto excluído que o controle guardou, visto do documento original: cada
 * ponto é uma posição do original onde o controlado tem `size` posições a mais.
 * É o mapa entre os dois — as posições do original, depois de cada passo, viram as
 * do controlado por aqui.
 */
class KeptContent {
  private points: { pos: number; size: number }[] = []

  /** A posição do original no controlado; no ponto exato, `assoc` diz de que lado. */
  map(pos: number, assoc: -1 | 1): number {
    let result = pos
    for (const point of this.points) {
      if (point.pos < pos || (point.pos === pos && assoc > 0)) result += point.size
    }
    return result
  }

  /** O mesmo mapa, como `Mapping`, para os passos que passam sem controle. */
  mapping(): Mapping {
    const mapping = new Mapping()
    let shift = 0
    for (const point of this.points) {
      mapping.appendMap(new StepMap([point.pos + shift, 0, point.size]))
      shift += point.size
    }
    return mapping
  }

  /**
   * Leva os pontos para depois do passo original. Os que caem em `absorbed` (o
   * trecho que o passo controlado guardou) viram um ponto só, no começo dele; os
   * que o passo apagou de fora a fora somem com ele.
   */
  advance(map: StepMap, absorbed?: { from: number; to: number; size: number }): void {
    const next: { pos: number; size: number }[] = []
    for (const point of this.points) {
      if (absorbed !== undefined && point.pos >= absorbed.from && point.pos <= absorbed.to) continue
      const result = map.mapResult(point.pos, -1)
      if (result.deletedAcross) continue
      next.push({ pos: result.pos, size: point.size })
    }
    if (absorbed !== undefined && absorbed.size > 0) next.push({ pos: absorbed.from, size: absorbed.size })
    next.sort((a, b) => a.pos - b.pos)
    this.points = []
    for (const point of next) {
      const last = this.points[this.points.length - 1]
      if (last !== undefined && last.pos === point.pos) last.size += point.size
      else this.points.push({ ...point })
    }
  }
}

// --- a exclusão ---------------------------------------------------------------

function hasMark(node: ProseMirrorNode, name: string): boolean {
  return node.marks.some((mark) => mark.type.name === name)
}

/** Texto que o próprio autor inseriu (e ninguém excluiu): apagar é apagar de verdade. */
function isOwnInsertion(node: ProseMirrorNode, author: string): boolean {
  return (
    !hasMark(node, DELETION) &&
    node.marks.some((mark) => mark.type.name === INSERTION && mark.attrs['author'] === author)
  )
}

function isOwnBlockInsertion(value: unknown, author: string): boolean {
  const revision = blockRevisionOf(value)
  return revision?.kind === 'ins' && revision.author === author
}

function blockRevision(kind: 'ins' | 'del', revision: Revision): Record<string, string> {
  return { kind, author: revision.author, date: revision.date }
}

/** O que fazer com um trecho excluído com o controle ligado. */
interface DeletionPlan {
  /** Trechos que ganham a marca `deletion`. */
  readonly marks: [number, number][]
  /** Trechos que saem de verdade: inserção própria, objeto de bloco, linha própria. */
  readonly drops: [number, number][]
  /** Blocos cuja marca de parágrafo passa a excluída. */
  readonly paragraphMarks: number[]
  /** Blocos cuja marca de parágrafo (inserida pelo autor) sai: juntam-se ao seguinte. */
  readonly joins: number[]
  /** Linhas de tabela que passam a excluídas. */
  readonly rows: number[]
}

/**
 * A marca de parágrafo de qual bloco a exclusão de [from, to] leva, se leva — a
 * do próprio bloco, quando o fim dele cai no meio do trecho; a do bloco de cima,
 * quando o trecho é o último bloco inteiro (o Backspace num parágrafo vazio).
 */
function paragraphMarkTaken(
  doc: ProseMirrorNode,
  node: ProseMirrorNode,
  pos: number,
  from: number,
  to: number,
): number | null {
  const end = pos + node.nodeSize
  if (end - 1 < from || end > to) return null
  if (end < to) return pos
  if (pos < from) return null
  const $pos = doc.resolve(pos)
  const index = $pos.index()
  if (index < $pos.parent.childCount - 1) return pos
  const previous = index > 0 ? $pos.parent.child(index - 1) : null
  if (previous?.isTextblock !== true) return null
  return pos - previous.nodeSize
}

/** Planeja a exclusão controlada de [from, to] — ou `null`, se ela não tem controle (coluna de tabela). */
function planDeletion(doc: ProseMirrorNode, from: number, to: number, author: string): DeletionPlan | null {
  const plan: DeletionPlan = { marks: [], drops: [], paragraphMarks: [], joins: [], rows: [] }
  const taken = new Set<number>()
  let untracked = false

  doc.nodesBetween(from, to, (node, pos) => {
    if (untracked) return false
    const end = pos + node.nodeSize
    const whole = pos >= from && end <= to

    if (node.isInline) {
      const start = Math.max(pos, from)
      const stop = Math.min(end, to)
      if (stop <= start || ZERO_WIDTH.has(node.type.name) || hasMark(node, DELETION)) return false
      if (isOwnInsertion(node, author)) plan.drops.push([start, stop])
      else plan.marks.push([start, stop])
      return false
    }

    if (node.isTextblock) {
      const owner = paragraphMarkTaken(doc, node, pos, from, to)
      if (owner !== null && !taken.has(owner)) {
        taken.add(owner)
        const block = doc.nodeAt(owner)!
        const after = owner + block.nodeSize
        const revision = blockRevisionOf(block.attrs['markRevision'])
        if (isOwnBlockInsertion(block.attrs['markRevision'], author) && canJoin(doc, after))
          plan.joins.push(owner)
        else if (revision?.kind !== 'del') plan.paragraphMarks.push(owner)
      }
      return true
    }

    if (node.type.name === 'tableRow' && whole) {
      const revision = blockRevisionOf(node.attrs['rowRevision'])
      if (isOwnBlockInsertion(node.attrs['rowRevision'], author)) plan.drops.push([pos, end])
      else if (revision?.kind !== 'del') plan.rows.push(pos)
      return false
    }

    // A célula inteira sem a linha inteira é coluna: o Word não a controla por
    // aqui, e a fase 1 já trava o documento com revisão de célula.
    if ((node.type.name === 'tableCell' || node.type.name === 'tableHeader') && whole) {
      untracked = true
      return false
    }

    if (node.isBlock && node.isLeaf && whole) {
      plan.drops.push([pos, end])
      return false
    }
    return true
  })

  return untracked ? null : plan
}

/** Aplica o plano. Devolve quantas posições saíram de verdade. */
function applyDeletion(tr: Transaction, plan: DeletionPlan, revision: Revision): number {
  const schema = tr.doc.type.schema
  const sizeBefore = tr.doc.content.size
  const mark = schema.marks[DELETION]!.create({ ...revision })

  for (const pos of plan.paragraphMarks) {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, markRevision: blockRevision('del', revision) })
  }
  for (const pos of plan.rows) {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, rowRevision: blockRevision('del', revision) })
  }
  for (const [from, to] of plan.marks) tr.addMark(from, to, mark)

  // O bloco que se junta fica com a marca de parágrafo do seguinte: é ela que
  // termina o parágrafo junto. De trás para frente — o que sai adiante não mexe
  // nas posições de antes, e na fila de junções cada bloco herda a marca que o
  // seguinte já herdou.
  const removals: { from: number; to: number; join?: number }[] = plan.drops.map(([from, to]) => ({
    from,
    to,
  }))
  for (const pos of plan.joins) {
    const end = pos + tr.doc.nodeAt(pos)!.nodeSize
    removals.push({ from: end - 1, to: end + 1, join: pos })
  }
  removals.sort((a, b) => b.from - a.from)
  for (const { from, to, join } of removals) {
    if (join !== undefined) {
      const node = tr.doc.nodeAt(join)!
      const next = tr.doc.nodeAt(join + node.nodeSize)
      tr.setNodeMarkup(join, undefined, { ...node.attrs, markRevision: next?.attrs['markRevision'] ?? null })
    }
    tr.delete(from, to)
  }

  return sizeBefore - tr.doc.content.size
}

// --- a inserção ---------------------------------------------------------------

/**
 * O trecho fechado só tem blocos de texto vazios — o parágrafo que o ProseMirror
 * põe no lugar do apagado (a célula limpa, o documento inteiro apagado). Aberto,
 * é o Enter, e esse entra.
 */
function onlyEmptyBlocks(slice: Slice): boolean {
  if (slice.openStart > 0) return false
  let empty = true
  slice.content.forEach((node) => {
    if (!node.isTextblock || node.content.size > 0) empty = false
  })
  return empty
}

/** Marca [from, to] como inserido pelo autor: o texto, as marcas de parágrafo e as linhas. */
function markInserted(tr: Transaction, from: number, to: number, revision: Revision): void {
  const schema = tr.doc.type.schema
  const insertion = schema.marks[INSERTION]!.create({ ...revision })
  const deletion = schema.marks[DELETION]!
  const inline: [number, number][] = []
  const paragraphs: number[] = []
  const rows: number[] = []

  tr.doc.nodesBetween(from, to, (node, pos) => {
    const end = pos + node.nodeSize
    if (node.isInline) {
      if (!ZERO_WIDTH.has(node.type.name)) inline.push([Math.max(pos, from), Math.min(end, to)])
      return false
    }
    if (node.type.name === 'tableRow' && pos >= from && end <= to) {
      rows.push(pos)
      return false
    }
    if (node.isTextblock && end - 1 >= from && end - 1 < to) paragraphs.push(pos)
    return true
  })

  for (const [start, stop] of inline) {
    if (stop <= start) continue
    tr.removeMark(start, stop, deletion)
    tr.addMark(start, stop, insertion)
  }
  for (const pos of paragraphs) {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, markRevision: blockRevision('ins', revision) })
  }
  for (const pos of rows) {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, rowRevision: blockRevision('ins', revision) })
  }
}

// --- a transação ---------------------------------------------------------------

/** O passo só põe ou tira pontas sem largura — comentário, marcador. */
function onlyAnchors(step: ReplaceStep, doc: ProseMirrorNode): boolean {
  const anchorsOnly = (fragment: Fragment): boolean => {
    let only = true
    fragment.descendants((node) => {
      if (!ZERO_WIDTH.has(node.type.name)) only = false
      return false
    })
    return only
  }
  return anchorsOnly(doc.slice(step.from, step.to).content) && anchorsOnly(step.slice.content)
}

/** Copia o que a transação traz além dos passos: metas, hora, marcas guardadas, rolagem. */
function copyExtras(from: Transaction, to: Transaction): void {
  const meta = (from as unknown as { meta: Record<string, unknown> }).meta
  for (const key of Object.keys(meta)) to.setMeta(key, meta[key])
  to.setTime(from.time)
  if (from.scrolledIntoView) to.scrollIntoView()
}

/**
 * Reescreve a transação com o controle de alterações: a mesma edição, com o que
 * sai guardado como excluído e o que entra marcado como inserido. Pura: recebe o
 * estado de antes e devolve outra transação sobre ele.
 */
export function trackTransaction(
  tr: Transaction,
  state: EditorState,
  author: string,
  now: Date,
  options: TrackOptions = {},
): Transaction {
  const revision: Revision = { author, date: revisionDate(now) }
  const tracked = state.tr
  const kept = new KeptContent()

  tr.steps.forEach((step: Step, index) => {
    const doc = tr.docs[index]!
    const map = step.getMap()

    if (!(step instanceof ReplaceStep) || onlyAnchors(step, doc)) {
      const mapped = step.map(kept.mapping())
      if (mapped !== null) tracked.maybeStep(mapped)
      kept.advance(map)
      return
    }

    const from = kept.map(step.from, -1)
    let to = kept.map(step.to, 1)

    // Na composição, o que sai sai de verdade; só o que entra é marcado.
    let plan: DeletionPlan | null = null
    if (step.to > step.from && options.composing !== true) {
      plan = planDeletion(tracked.doc, from, to, author)
      if (plan === null) {
        const mapped = step.map(kept.mapping())
        if (mapped !== null) tracked.maybeStep(mapped)
        kept.advance(map)
        return
      }
    }

    let removed = 0
    if (plan !== null) {
      removed = applyDeletion(tracked, plan, revision)
    } else if (step.to > step.from) {
      const sizeBefore = tracked.doc.content.size
      tracked.delete(from, to)
      removed = sizeBefore - tracked.doc.content.size
    }
    to -= removed

    // O que entra vai depois do que ficou excluído, como no Word.
    const slice = step.slice
    if (slice.size > 0 && !(step.to > step.from && onlyEmptyBlocks(slice))) {
      const before = tracked.steps.length
      // O bloco que o Enter abre leva a marca de parágrafo de onde ele caiu — a do
      // documento controlado, e não a do original, que o trecho copiou.
      const $at = tracked.doc.resolve(to)
      const tail: unknown = $at.parent.isTextblock ? $at.parent.attrs['markRevision'] : undefined
      tracked.replace(to, to, slice)
      if (tracked.steps.length > before) {
        const mapping = tracked.mapping.slice(before)
        const end = mapping.map(to, 1)
        markInserted(tracked, mapping.map(to, -1), end, revision)
        const $end = tracked.doc.resolve(end)
        if (slice.openEnd > 0 && tail !== undefined && $end.parent.isTextblock && $end.depth > 0) {
          const pos = $end.before()
          if ($end.parent.attrs['markRevision'] !== tail) {
            tracked.setNodeMarkup(pos, undefined, { ...$end.parent.attrs, markRevision: tail })
          }
        }
      }
    }

    kept.advance(map, { from: step.from, to: step.to, size: to - from })
  })

  copyExtras(tr, tracked)

  // O Backspace deixa o cursor antes do que ficou excluído; o Delete e o resto, depois.
  const before = state.selection
  const after = tr.selection
  const backward = before.empty && after.empty && after.head < before.head
  const assoc = backward ? -1 : 1
  if (after instanceof TextSelection) {
    const doc = tracked.doc
    const anchor = Math.min(kept.map(after.anchor, assoc), doc.content.size)
    const head = Math.min(kept.map(after.head, assoc), doc.content.size)
    tracked.setSelection(TextSelection.between(doc.resolve(anchor), doc.resolve(head)))
  } else {
    tracked.setSelection(after.map(tracked.doc, kept.mapping()))
  }
  if (tr.storedMarksSet) tracked.setStoredMarks(tr.storedMarks)

  return tracked
}

// --- o desfazer ---------------------------------------------------------------

/** O grupo do desfazer em que as últimas edições controladas caíram. */
export interface TrackGroup {
  readonly id: string
  /** A hora da última edição do grupo. */
  readonly time: number
  /** Onde o cursor ficou depois dela. */
  readonly head: number
}

/** O mesmo intervalo do `prosemirror-history` (`newGroupDelay`). */
const GROUP_DELAY_MS = 500
let groupCount = 0

/**
 * Junta a edição controlada às anteriores no desfazer, como o histórico junta as
 * sem controle.
 *
 * O histórico agrupa pelo relógio e pela vizinhança dos trechos mudados, mas a
 * exclusão controlada não muda trecho nenhum — só põe a marca — e cada Backspace
 * virava um passo próprio. A vizinhança aqui se mede pela edição pedida
 * (`original`): ela começa onde o cursor ficou depois da anterior? E dentro do
 * mesmo intervalo? Então vai junto, pela meta `composition`, que é a que o
 * histórico usa para não partir a composição do IME — um valor nosso, que não
 * colide com os números dela. Devolve o grupo de agora.
 */
export function joinHistoryGroup(
  original: Transaction,
  tracked: Transaction,
  previous: TrackGroup | null,
): TrackGroup | null {
  // A composição de verdade manda no próprio grupo.
  if (original.getMeta('composition') !== undefined) return null
  const map = original.mapping.maps[0]
  let adjacent = false
  if (previous !== null && original.time - previous.time < GROUP_DELAY_MS && map !== undefined) {
    map.forEach((start, end) => {
      if (start <= previous.head && end >= previous.head) adjacent = true
    })
  }
  const id = adjacent && previous !== null ? previous.id : `track-${++groupCount}`
  tracked.setMeta('composition', id)
  return { id, time: original.time, head: tracked.selection.head }
}

// --- a área de transferência ---------------------------------------------------

function isRevisionMark(mark: Mark): boolean {
  return mark.type.name === INSERTION || mark.type.name === DELETION
}

/** O conteúdo sem marcas nem atributos de revisão: o que se cola é texto novo. */
function withoutRevisions(fragment: Fragment): Fragment {
  const children: ProseMirrorNode[] = []
  fragment.forEach((node) => {
    if (node.isText) {
      children.push(node.mark(node.marks.filter((mark) => !isRevisionMark(mark))))
      return
    }
    const attrs =
      'markRevision' in node.attrs || 'rowRevision' in node.attrs
        ? {
            ...node.attrs,
            ...('markRevision' in node.attrs ? { markRevision: null } : {}),
            ...('rowRevision' in node.attrs ? { rowRevision: null } : {}),
          }
        : node.attrs
    children.push(
      node.type.create(
        attrs,
        withoutRevisions(node.content),
        node.marks.filter((mark) => !isRevisionMark(mark)),
      ),
    )
  })
  return Fragment.fromArray(children)
}

/** O conteúdo sem o que está excluído: copiar leva o texto como ele fica. */
function withoutDeleted(fragment: Fragment): Fragment {
  const children: ProseMirrorNode[] = []
  fragment.forEach((node) => {
    if (node.isInline && hasMark(node, DELETION)) return
    children.push(node.isLeaf ? node : node.copy(withoutDeleted(node.content)))
  })
  return Fragment.fromArray(children)
}

export function stripRevisions(slice: Slice): Slice {
  return new Slice(withoutRevisions(slice.content), slice.openStart, slice.openEnd)
}

export function stripDeleted(slice: Slice): Slice {
  return new Slice(withoutDeleted(slice.content), slice.openStart, slice.openEnd)
}

// --- a extensão ---------------------------------------------------------------

export interface TrackInputOptions {
  /** O controle está ligado? Consultado a cada transação — muda com o editor no ar. */
  readonly isTracking: () => boolean
  /** Quem assina as alterações. */
  readonly author: () => string
}

const trackInputKey = new PluginKey('trackInput')

/** Se o bloco tem algum trecho com marca de revisão. */
function hasRevisionInside(block: ProseMirrorNode): boolean {
  let found = false
  block.forEach((child) => {
    if (child.marks.some((mark) => mark.type.name === INSERTION || mark.type.name === DELETION)) found = true
  })
  return found
}

/**
 * O trecho que o Ctrl+Backspace (ou Ctrl+Delete) apaga, em posições dentro do
 * bloco: até o começo da palavra anterior (ou o começo da seguinte). O já
 * excluído e as âncoras sem largura são transparentes — a palavra é a do texto
 * que se vê. `null` quando não há o que apagar (a ponta do bloco).
 */
export function wordRangeAt(
  block: ProseMirrorNode,
  offset: number,
  backward: boolean,
): [number, number] | null {
  // Uma entrada por posição do bloco: 'w' letra, 's' espaço, 't' transparente.
  const kinds: string[] = []
  block.forEach((child) => {
    if (child.isText) {
      const deleted = hasMark(child, DELETION)
      for (const char of child.text ?? '') {
        const kind = deleted ? 't' : /\s/.test(char) ? 's' : 'w'
        for (let unit = 0; unit < char.length; unit++) kinds.push(kind)
      }
    } else {
      const kind = ZERO_WIDTH.has(child.type.name) ? 't' : 'w'
      for (let unit = 0; unit < child.nodeSize; unit++) kinds.push(kind)
    }
  })

  let index = offset
  if (backward) {
    while (index > 0 && kinds[index - 1] !== 'w') index--
    while (index > 0 && kinds[index - 1] !== 's') index--
    return index === offset ? null : [index, offset]
  }
  while (index < kinds.length && kinds[index] !== 's') index++
  while (index < kinds.length && kinds[index] !== 'w') index++
  return index === offset ? null : [offset, index]
}

export const TrackInput = Extension.create<TrackInputOptions>({
  name: 'trackInput',

  addOptions() {
    return { isTracking: () => false, author: () => '' }
  },

  addStorage() {
    return { group: null as TrackGroup | null }
  },

  dispatchTransaction({ transaction, next }) {
    if (!this.options.isTracking() || !shouldTrack(transaction)) {
      // Outra edição no meio: a seguinte controlada começa grupo novo.
      if (transaction.docChanged) this.storage.group = null
      next(transaction)
      return
    }
    const view = this.editor.view
    let tracked: Transaction
    try {
      tracked = trackTransaction(transaction, view.state, this.options.author(), new Date(), {
        composing: view.composing,
      })
    } catch (error) {
      // Melhor a edição sem controle que a edição perdida.
      console.error(error)
      tracked = transaction
    }
    if (tracked !== transaction)
      this.storage.group = joinHistoryGroup(transaction, tracked, this.storage.group)
    next(tracked)
  },

  addProseMirrorPlugins() {
    const options = this.options
    return [
      new Plugin({
        key: trackInputKey,
        props: {
          // O Backspace e o Delete de um caractere, com o controle ligado, feitos
          // aqui e não pelo navegador: apagando sozinho ele mexe também no trecho
          // excluído vizinho (o espaço do começo vira `&nbsp;`), e a releitura da
          // tela devolvia esse trecho como texto novo — inserido de volta.
          handleKeyDown(view, event) {
            if (view.composing) return false
            if (event.key !== 'Backspace' && event.key !== 'Delete') return false
            if (event.metaKey || event.altKey || event.shiftKey) return false
            const { selection } = view.state
            if (!selection.empty) return false
            const $cursor = selection.$head
            const backward = event.key === 'Backspace'

            // Ctrl: a palavra. O navegador, apagando sozinho, refaz o `<del>`
            // vizinho como tachado comum — a revisão virava formatação mesmo com o
            // controle desligado. Então, perto de revisão ou com o controle ligado,
            // a palavra sai por aqui.
            if (event.ctrlKey) {
              if (!options.isTracking() && !hasRevisionInside($cursor.parent)) return false
              const range = wordRangeAt($cursor.parent, $cursor.parentOffset, backward)
              if (range === null) return false
              const start = $cursor.start()
              view.dispatch(view.state.tr.delete(start + range[0], start + range[1]).scrollIntoView())
              return true
            }
            if (!options.isTracking()) return false
            const node = backward ? $cursor.nodeBefore : $cursor.nodeAfter
            if (node === null || !node.isText || node.text === undefined) return false
            // Um caractere, inteiro: o par substituto de um emoji vai junto.
            const text = node.text
            const unit = backward ? text.charCodeAt(text.length - 1) : text.charCodeAt(0)
            const surrogate = backward ? unit >= 0xdc00 && unit <= 0xdfff : unit >= 0xd800 && unit <= 0xdbff
            const size = surrogate && text.length > 1 ? 2 : 1
            const from = backward ? $cursor.pos - size : $cursor.pos
            view.dispatch(view.state.tr.delete(from, from + size).scrollIntoView())
            return true
          },
          // O que se cola (e se arrasta) entra como texto novo; com o controle
          // ligado, é ele que vira inserção.
          transformPasted: (slice) => stripRevisions(slice),
          // O excluído não vai para a área de transferência.
          transformCopied: (slice) => stripDeleted(slice),
          // E nem do texto puro, que o Tiptap tira da seleção sem passar acima.
          clipboardTextSerializer: (slice) => {
            const content = stripDeleted(slice).content
            return content.textBetween(0, content.size, '\n\n', (leaf) =>
              leaf.type.name === 'hardBreak' ? '\n' : '',
            )
          },
        },
      }),
    ]
  },
})
