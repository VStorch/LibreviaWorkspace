import { Extension } from '@tiptap/core'
import { Fragment, Slice, type Mark, type Node as ProseMirrorNode, type ResolvedPos } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Mapping, ReplaceStep, StepMap, canJoin, type Step } from '@tiptap/pm/transform'
import type { EditorView } from '@tiptap/pm/view'
import { DELETION, INSERTION, ZERO_WIDTH, blockRevisionOf, characterSize } from './track-changes.js'

/**
 * A transação é reescrita **antes** de ser aplicada (`dispatchTransaction`): uma
 * só, com a seleção certa e um passo de desfazer. Em cada `ReplaceStep`:
 *
 * - o que entra ganha `insertion` do autor e perde `deletion`;
 * - o que sai **fica**, com `deletion` — menos a inserção do próprio autor, que
 *   sai de verdade como no Word, e o já excluído, que fica como estava;
 * - o Enter que sai vira `markRevision del` do bloco de cima, e o que entra,
 *   `markRevision ins`;
 * - a linha de tabela inteira ganha `rowRevision`.
 *
 * Formatação passa sem controle, assim como desfazer, aceitar e rejeitar
 * (`SKIP_TRACKING`) e a ponta sozinha de comentário ou marcador. Na composição do
 * IME só a inserção é marcada: devolver o que ela apaga desmontaria o DOM dela.
 */

/** Aceitar e rejeitar, por exemplo. */
export const SKIP_TRACKING = 'trackChanges:skip'

/** Desfazer e refazer não são edição nova. */
const HISTORY_META = 'history$'

/** Ao minuto, como o Word: o que se digita no mesmo minuto se funde num trecho só. */
export function revisionDate(now: Date): string {
  return `${now.toISOString().slice(0, 16)}:00Z`
}

interface Revision {
  readonly author: string
  readonly date: string
}

export interface TrackOptions {
  readonly composing?: boolean
}

export function shouldTrack(tr: Transaction): boolean {
  if (!tr.docChanged) return false
  if (tr.getMeta(SKIP_TRACKING) === true) return false
  if (tr.getMeta(HISTORY_META) !== undefined) return false
  return tr.getMeta('addToHistory') !== false
}

/**
 * O texto excluído que o controle guardou, visto do original: cada ponto é uma
 * posição do original onde o controlado tem `size` posições a mais.
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

  mapping(): Mapping {
    const mapping = new Mapping()
    let shift = 0
    for (const point of this.points) {
      mapping.appendMap(new StepMap([point.pos + shift, 0, point.size]))
      shift += point.size
    }
    return mapping
  }

  /** Os pontos dentro de `absorbed` viram um só, no começo dele; os que o passo apagou somem. */
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

function hasMark(node: ProseMirrorNode, name: string): boolean {
  return node.marks.some((mark) => mark.type.name === name)
}

/** Apagar o que o próprio autor inseriu é apagar de verdade. */
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

interface DeletionPlan {
  readonly marks: [number, number][]
  /** Trechos que saem de verdade: inserção própria, objeto de bloco, linha própria. */
  readonly drops: [number, number][]
  /** Blocos cuja marca de parágrafo passa a excluída. */
  readonly paragraphMarks: number[]
  /** Blocos cuja marca de parágrafo (inserida pelo autor) sai: juntam-se ao seguinte. */
  readonly joins: number[]
  readonly rows: number[]
}

/** A do próprio bloco, quando o fim dele cai no trecho; a do de cima, no Backspace num parágrafo vazio. */
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

/** `null` quando a exclusão não tem controle, como numa coluna de tabela. */
function planDeletion(doc: ProseMirrorNode, from: number, to: number, author: string): DeletionPlan | null {
  const planner = new DeletionPlanner(doc, from, to, author)
  doc.nodesBetween(from, to, (node, pos) => planner.visit(node, pos))
  return planner.untracked ? null : planner.plan
}

class DeletionPlanner {
  readonly plan: DeletionPlan = { marks: [], drops: [], paragraphMarks: [], joins: [], rows: [] }
  untracked = false
  private readonly taken = new Set<number>()

  constructor(
    private readonly doc: ProseMirrorNode,
    private readonly from: number,
    private readonly to: number,
    private readonly author: string,
  ) {}

  /** O retorno é o do `nodesBetween`: descer ou não aos filhos. */
  visit(node: ProseMirrorNode, pos: number): boolean {
    if (this.untracked) return false
    const end = pos + node.nodeSize

    if (node.isInline) {
      this.inline(node, pos, end)
      return false
    }

    if (node.isTextblock) {
      this.paragraphMark(node, pos)
      return true
    }

    return pos >= this.from && end <= this.to ? this.wholeBlock(node, pos, end) : true
  }

  private wholeBlock(node: ProseMirrorNode, pos: number, end: number): boolean {
    if (node.type.name === 'tableRow') {
      const revision = blockRevisionOf(node.attrs['rowRevision'])
      if (isOwnBlockInsertion(node.attrs['rowRevision'], this.author)) this.plan.drops.push([pos, end])
      else if (revision?.kind !== 'del') this.plan.rows.push(pos)
      return false
    }

    // A célula inteira sem a linha inteira é coluna: o Word não a controla por
    // aqui, e o documento com revisão de célula já abre travado.
    if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
      this.untracked = true
      return false
    }

    if (node.isBlock && node.isLeaf) {
      this.plan.drops.push([pos, end])
      return false
    }
    return true
  }

  private inline(node: ProseMirrorNode, pos: number, end: number): void {
    const start = Math.max(pos, this.from)
    const stop = Math.min(end, this.to)
    if (stop <= start || ZERO_WIDTH.has(node.type.name) || hasMark(node, DELETION)) return
    if (isOwnInsertion(node, this.author)) this.plan.drops.push([start, stop])
    else this.plan.marks.push([start, stop])
  }

  private paragraphMark(node: ProseMirrorNode, pos: number): void {
    const owner = paragraphMarkTaken(this.doc, node, pos, this.from, this.to)
    if (owner === null || this.taken.has(owner)) return
    this.taken.add(owner)
    const block = this.doc.nodeAt(owner)!
    const after = owner + block.nodeSize
    const revision = blockRevisionOf(block.attrs['markRevision'])
    if (isOwnBlockInsertion(block.attrs['markRevision'], this.author) && canJoin(this.doc, after))
      this.plan.joins.push(owner)
    else if (revision?.kind !== 'del') this.plan.paragraphMarks.push(owner)
  }
}

/** Devolve quantas posições saíram de verdade. */
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

  // O bloco que se junta fica com a marca de parágrafo do seguinte. De trás para
  // frente, para o que sai adiante não mexer nas posições de antes.
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

/** O parágrafo que o ProseMirror põe no lugar do apagado; aberto, é o Enter, e esse entra. */
function onlyEmptyBlocks(slice: Slice): boolean {
  if (slice.openStart > 0) return false
  let empty = true
  slice.content.forEach((node) => {
    if (!node.isTextblock || node.content.size > 0) empty = false
  })
  return empty
}

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

function copyExtras(from: Transaction, to: Transaction): void {
  const meta = (from as unknown as { meta: Record<string, unknown> }).meta
  for (const key of Object.keys(meta)) to.setMeta(key, meta[key])
  to.setTime(from.time)
  if (from.scrolledIntoView) to.scrollIntoView()
}

/** Devolve quanto o documento encolheu. */
function deleteForReal(tracked: Transaction, from: number, to: number): number {
  const sizeBefore = tracked.doc.content.size
  tracked.delete(from, to)
  return sizeBefore - tracked.doc.content.size
}

function insertTracked(tracked: Transaction, to: number, slice: Slice, revision: Revision): void {
  const before = tracked.steps.length
  // A marca de parágrafo do documento controlado, e não a do original.
  const $at = tracked.doc.resolve(to)
  const tail: unknown = $at.parent.isTextblock ? $at.parent.attrs['markRevision'] : undefined
  tracked.replace(to, to, slice)
  if (tracked.steps.length === before) return

  const mapping = tracked.mapping.slice(before)
  const end = mapping.map(to, 1)
  markInserted(tracked, mapping.map(to, -1), end, revision)
  const $end = tracked.doc.resolve(end)
  if (slice.openEnd > 0 && tail !== undefined && $end.parent.isTextblock && $end.depth > 0) {
    if ($end.parent.attrs['markRevision'] !== tail) {
      tracked.setNodeMarkup($end.before(), undefined, { ...$end.parent.attrs, markRevision: tail })
    }
  }
}

/** Pura: recebe o estado de antes e devolve outra transação sobre ele. */
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
    const passThrough = (): void => {
      const mapped = step.map(kept.mapping())
      if (mapped !== null) tracked.maybeStep(mapped)
      kept.advance(map)
    }

    if (!(step instanceof ReplaceStep) || onlyAnchors(step, doc)) return passThrough()

    const from = kept.map(step.from, -1)
    let to = kept.map(step.to, 1)
    const deletes = step.to > step.from

    // Na composição, o que sai sai de verdade; só o que entra é marcado.
    const tracksDeletion = deletes && options.composing !== true
    const plan = tracksDeletion ? planDeletion(tracked.doc, from, to, author) : null
    if (tracksDeletion && plan === null) return passThrough()

    if (plan !== null) to -= applyDeletion(tracked, plan, revision)
    else if (deletes) to -= deleteForReal(tracked, from, to)

    // O que entra vai depois do que ficou excluído, como no Word.
    if (step.slice.size > 0 && !(deletes && onlyEmptyBlocks(step.slice))) {
      insertTracked(tracked, to, step.slice, revision)
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

export interface TrackGroup {
  readonly id: string
  readonly time: number
  readonly head: number
}

/** O de `newGroupDelay` do `prosemirror-history`. */
const GROUP_DELAY_MS = 500
let groupCount = 0

/**
 * A exclusão controlada só põe a marca, e o histórico, que agrupa pela
 * vizinhança dos trechos mudados, faria de cada Backspace um passo. Aqui a
 * vizinhança é a da edição pedida, e o agrupamento vai pela meta `composition`,
 * a mesma que o histórico usa para o IME.
 */
export function joinHistoryGroup(
  original: Transaction,
  tracked: Transaction,
  previous: TrackGroup | null,
): TrackGroup | null {
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

function isRevisionMark(mark: Mark): boolean {
  return mark.type.name === INSERTION || mark.type.name === DELETION
}

/** O que se cola é texto novo. */
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

/** Copiar leva o texto como ele fica. */
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

export interface TrackInputOptions {
  /** Consultado a cada transação: muda com o editor no ar. */
  readonly isTracking: () => boolean
  readonly author: () => string
}

const trackInputKey = new PluginKey('trackInput')

function hasRevisionInside(block: ProseMirrorNode): boolean {
  let found = false
  block.forEach((child) => {
    if (child.marks.some((mark) => mark.type.name === INSERTION || mark.type.name === DELETION)) found = true
  })
  return found
}

/** O já excluído e as âncoras sem largura são transparentes; `null` na ponta do bloco. */
export function wordRangeAt(
  block: ProseMirrorNode,
  offset: number,
  backward: boolean,
): [number, number] | null {
  // 'w' letra, 's' espaço, 't' transparente.
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

/** A referência de nota só como a marca: `textBetween` despejaria o corpo no meio da frase. */
function plainTextOf(fragment: Fragment): string {
  const blocks: string[] = []
  let inline = ''
  fragment.forEach((node) => {
    if (node.isText) inline += node.text ?? ''
    else if (node.type.name === 'hardBreak') inline += '\n'
    else if (node.isInline) inline += node.type.name === 'noteRef' ? String(node.attrs['mark'] ?? '') : ''
    else blocks.push(plainTextOf(node.content))
  })
  if (inline !== '') blocks.unshift(inline)
  return blocks.join('\n\n')
}

/**
 * Feitos aqui, e não pelo navegador, que mexeria no trecho excluído vizinho e o
 * devolveria como texto novo. Exportado para o corpo da nota, que é outro
 * `EditorView`, sem os plugins do editor.
 */
export function trackedDeleteKey(view: EditorView, event: KeyboardEvent, isTracking: () => boolean): boolean {
  if (view.composing) return false
  if (event.key !== 'Backspace' && event.key !== 'Delete') return false
  if (event.metaKey || event.altKey || event.shiftKey) return false
  const { selection } = view.state
  if (!selection.empty) return false
  const backward = event.key === 'Backspace'

  // A palavra também: o navegador refaria o `<del>` vizinho como tachado comum.
  if (event.ctrlKey) return deleteWord(view, selection.$head, backward, isTracking)
  if (!isTracking()) return false
  return deleteCharacter(view, selection.$head, backward)
}

function deleteWord(
  view: EditorView,
  $cursor: ResolvedPos,
  backward: boolean,
  isTracking: () => boolean,
): boolean {
  if (!isTracking() && !hasRevisionInside($cursor.parent)) return false
  const range = wordRangeAt($cursor.parent, $cursor.parentOffset, backward)
  if (range === null) return false
  const start = $cursor.start()
  view.dispatch(view.state.tr.delete(start + range[0], start + range[1]).scrollIntoView())
  return true
}

function deleteCharacter(view: EditorView, $cursor: ResolvedPos, backward: boolean): boolean {
  const node = backward ? $cursor.nodeBefore : $cursor.nodeAfter
  if (node === null || !node.isText || node.text === undefined) return false
  const size = characterSize(node.text, backward)
  const from = backward ? $cursor.pos - size : $cursor.pos
  view.dispatch(view.state.tr.delete(from, from + size).scrollIntoView())
  return true
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
          handleKeyDown: (view, event) => trackedDeleteKey(view, event, options.isTracking),
          // O que se cola entra como texto novo; o excluído não vai para a área
          // de transferência, nem no texto puro.
          transformPasted: (slice) => stripRevisions(slice),
          transformCopied: (slice) => stripDeleted(slice),
          clipboardTextSerializer: (slice) => {
            return plainTextOf(stripDeleted(slice).content)
          },
        },
      }),
    ]
  },
})
