import { Extension, Mark, mergeAttributes } from '@tiptap/core'
import type { Mark as ProseMirrorMark, Node as ProseMirrorNode } from '@tiptap/pm/model'
import { TextSelection, type Transaction } from '@tiptap/pm/state'
import { canJoin, type Mapping } from '@tiptap/pm/transform'
import { CharacterCount } from '@tiptap/extensions'

/**
 * A revisão de texto é marca do trecho (`insertion`, `deletion`), com autor,
 * data e o `w:id` original (`rid`); as duas não se excluem: o `w:ins` que embrulha
 * um `w:del` é o texto que alguém inseriu e outro excluiu. A movimentação leva
 * `move` e `moveName` para voltar como era. Enter e linha de tabela são atributos
 * do bloco (`markRevision`, `rowRevision`), como no arquivo.
 */

export const INSERTION = 'insertion'
export const DELETION = 'deletion'

/** As pontas sem largura de marcador e de comentário: não partem uma alteração. */
export const ZERO_WIDTH: ReadonlySet<string> = new Set([
  'bookmarkStart',
  'bookmarkEnd',
  'commentStart',
  'commentEnd',
])

export interface BlockRevision {
  readonly kind: 'ins' | 'del'
  readonly author?: string
  readonly date?: string
  readonly rid?: string
}

export type ChangeKind =
  'insertion' | 'deletion' | 'markInsertion' | 'markDeletion' | 'rowInsertion' | 'rowDeletion'

export interface RevisionChange {
  readonly kind: ChangeKind
  readonly from: number
  readonly to: number
  readonly author: string | null
  readonly date: string | null
  /** Sem as pontas sem largura entre os pedaços. */
  readonly segments: readonly { readonly from: number; readonly to: number }[]
  /** Só nas alterações de texto. */
  readonly mark?: ProseMirrorMark
}

function attrString(value: unknown): string | null {
  return typeof value === 'string' && value !== '' ? value : null
}

export function blockRevisionOf(value: unknown): BlockRevision | null {
  if (typeof value !== 'object' || value === null) return null
  const kind = (value as { kind?: unknown }).kind
  return kind === 'ins' || kind === 'del' ? (value as BlockRevision) : null
}

/** Mesmo autor e mesmo `rid` (o arquivo), ou mesmo autor sem `rid` (o editor). */
function sameRevision(a: ProseMirrorMark, b: ProseMirrorMark): boolean {
  return a.attrs['author'] === b.attrs['author'] && a.attrs['rid'] === b.attrs['rid']
}

/** Na ordem do texto. As pontas de marcador e de comentário no meio não partem o trecho e sobrevivem ao aceite. */
export function revisionChangesOf(doc: ProseMirrorNode): RevisionChange[] {
  const changes: RevisionChange[] = []
  collectChanges(doc, 0, changes)
  return changes.sort((a, b) => a.from - b.from || a.to - b.to)
}

/** Desce nos corpos de nota: aceitar todas não pode deixá-las para trás. */
function collectChanges(parent: ProseMirrorNode, base: number, changes: RevisionChange[]): void {
  parent.descendants((node, relative) => {
    const pos = base + relative
    if (node.type.name === 'tableRow') {
      const revision = blockRevisionOf(node.attrs['rowRevision'])
      if (revision !== null) {
        changes.push({
          kind: revision.kind === 'ins' ? 'rowInsertion' : 'rowDeletion',
          from: pos,
          to: pos + node.nodeSize,
          author: attrString(revision.author),
          date: attrString(revision.date),
          segments: [],
        })
      }
      return true
    }

    if (!node.isTextblock) return true

    for (const type of [INSERTION, DELETION] as const) {
      let open: { mark: ProseMirrorMark; segments: { from: number; to: number }[] } | null = null
      const close = (): void => {
        if (open === null) return
        const first = open.segments[0]!
        const last = open.segments[open.segments.length - 1]!
        changes.push({
          kind: type,
          from: first.from,
          to: last.to,
          author: attrString(open.mark.attrs['author']),
          date: attrString(open.mark.attrs['date']),
          segments: open.segments,
          mark: open.mark,
        })
        open = null
      }

      node.forEach((child, offset) => {
        const from = pos + 1 + offset
        if (ZERO_WIDTH.has(child.type.name)) return
        const mark = child.marks.find((candidate) => candidate.type.name === type)
        if (mark === undefined) {
          close()
          return
        }
        if (open !== null && !sameRevision(open.mark, mark)) close()
        if (open === null) open = { mark, segments: [] }
        const previous = open.segments[open.segments.length - 1]
        if (previous !== undefined && previous.to === from) previous.to = from + child.nodeSize
        else open.segments.push({ from, to: from + child.nodeSize })
      })
      close()
    }

    const revision = blockRevisionOf(node.attrs['markRevision'])
    if (revision !== null) {
      const end = pos + node.nodeSize - 1
      changes.push({
        kind: revision.kind === 'ins' ? 'markInsertion' : 'markDeletion',
        from: end,
        to: end,
        author: attrString(revision.author),
        date: attrString(revision.date),
        segments: [],
      })
    }

    node.forEach((child, offset) => {
      if (child.type.name === 'noteRef') collectChanges(child, pos + 1 + offset + 1, changes)
    })
    return false
  })
}

/** O trecho que contém o cursor, senão a marca do parágrafo, senão a linha. */
export function changeAt(doc: ProseMirrorNode, pos: number): RevisionChange | null {
  const changes = revisionChangesOf(doc)
  const inline = changes.find(
    (change) =>
      (change.kind === INSERTION || change.kind === DELETION) && change.from <= pos && pos <= change.to,
  )
  if (inline !== undefined) return inline

  const $pos = doc.resolve(pos)
  for (let depth = $pos.depth; depth > 0; depth--) {
    const node = $pos.node(depth)
    const start = $pos.before(depth)
    if (node.isTextblock) {
      const mark = changes.find(
        (change) =>
          (change.kind === 'markInsertion' || change.kind === 'markDeletion') &&
          change.from === start + node.nodeSize - 1,
      )
      if (mark !== undefined) return mark
    }
    if (node.type.name === 'tableRow') {
      const row = changes.find(
        (change) =>
          (change.kind === 'rowInsertion' || change.kind === 'rowDeletion') && change.from === start,
      )
      if (row !== undefined) return row
    }
  }
  return null
}

/** Rejeitar faz o contrário. */
type Effect = 'keep' | 'drop'

function effectOf(kind: ChangeKind, accept: boolean): Effect {
  const inserted = kind === INSERTION || kind === 'markInsertion' || kind === 'rowInsertion'
  return inserted === accept ? 'keep' : 'drop'
}

/** As posições são as do documento em que a alteração foi achada, mapeadas desde `since`. */
function settle(tr: Transaction, change: RevisionChange, accept: boolean, since: number): void {
  SETTLERS[change.kind](tr, change, effectOf(change.kind, accept), tr.mapping.slice(since))
}

type Settler = (tr: Transaction, change: RevisionChange, effect: Effect, mapping: Mapping) => void

const settleText: Settler = (tr, change, effect, mapping) => {
  const segments = change.segments
    .map((segment) => ({ from: mapping.map(segment.from, 1), to: mapping.map(segment.to, -1) }))
    .filter((segment) => segment.to > segment.from)
  for (const segment of [...segments].reverse()) {
    if (effect === 'drop') tr.delete(segment.from, segment.to)
    else tr.removeMark(segment.from, segment.to, change.mark)
  }
}

const settleParagraphMark: Settler = (tr, change, effect, mapping) => {
  const end = mapping.map(change.from)
  const paragraph = tr.doc.resolve(end).parent
  const start = end - paragraph.nodeSize + 1
  if (start < 0 || tr.doc.nodeAt(start) !== paragraph) return
  tr.setNodeMarkup(start, undefined, { ...paragraph.attrs, markRevision: null })
  // O Enter que sai junta os parágrafos, como no Word.
  const after = start + paragraph.nodeSize
  if (effect === 'drop' && after < tr.doc.content.size && canJoin(tr.doc, after)) tr.join(after)
}

const settleRow: Settler = (tr, change, effect, mapping) => {
  const from = mapping.map(change.from, 1)
  const row = tr.doc.nodeAt(from)
  if (row === null || row.type.name !== 'tableRow') return
  if (effect === 'keep') {
    tr.setNodeMarkup(from, undefined, { ...row.attrs, rowRevision: null })
    return
  }
  const $row = tr.doc.resolve(from)
  const table = $row.parent
  // Tabela sem linha não existe.
  if (table.childCount === 1) {
    const tableStart = $row.before()
    tr.delete(tableStart, tableStart + table.nodeSize)
  } else {
    tr.delete(from, from + row.nodeSize)
  }
}

const SETTLERS: Record<ChangeKind, Settler> = {
  insertion: settleText,
  deletion: settleText,
  markInsertion: settleParagraphMark,
  markDeletion: settleParagraphMark,
  rowInsertion: settleRow,
  rowDeletion: settleRow,
}

/** Devolve se havia alguma. */
export function settleChangeAt(tr: Transaction, pos: number, accept: boolean): boolean {
  const change = changeAt(tr.doc, pos)
  if (change === null) return false
  settle(tr, change, accept, tr.steps.length)
  return true
}

/** Numa transação só: um desfazer devolve tudo. */
export function settleAllChanges(tr: Transaction, accept: boolean): boolean {
  const changes = revisionChangesOf(tr.doc)
  if (changes.length === 0) return false
  const since = tr.steps.length
  for (const change of [...changes].reverse()) settle(tr, change, accept, since)
  return true
}

export function adjacentChange(doc: ProseMirrorNode, pos: number, direction: 1 | -1): RevisionChange | null {
  const changes = revisionChangesOf(doc)
  // Para frente, a que começa no cursor vale, menos a que acaba nele (a recém-escolhida).
  if (direction === 1) {
    return changes.find((change) => change.from > pos || (change.from === pos && change.to > pos)) ?? null
  }
  return [...changes].reverse().find((change) => change.from < pos && change.to <= pos) ?? null
}

export function selectChange(tr: Transaction, change: RevisionChange): Transaction {
  const doc = tr.doc
  const selection =
    change.kind === INSERTION || change.kind === DELETION
      ? TextSelection.create(doc, change.from, change.to)
      : change.kind === 'markInsertion' || change.kind === 'markDeletion'
        ? TextSelection.create(doc, change.from)
        : TextSelection.near(doc.resolve(change.from + 1))
  return tr.setSelection(selection).scrollIntoView()
}

function isDeleted(node: ProseMirrorNode): boolean {
  return node.marks.some((mark) => mark.type.name === DELETION)
}

/** Com `hide`, o excluído vira esse caractere repetido: a busca precisa das posições, e ele nunca casa. */
export function textWithoutDeletions(
  node: ProseMirrorNode,
  blockSeparator: string | undefined,
  leafText: string | ((leaf: ProseMirrorNode) => string),
  hide?: string,
): string {
  let text = ''
  let first = true
  node.descendants((child) => {
    // A nota é um nó só no texto do parágrafo; na contagem ela entra, como no Word.
    if (child.type.name === 'noteRef') {
      text += noteText(child, blockSeparator, leafText, hide)
      return false
    }
    const own = ownText(child, leafText, hide)
    const separated = child.isTextblock || (child.isLeaf && own !== '')
    if (child.isBlock && separated && blockSeparator !== undefined) {
      if (first) first = false
      else text += blockSeparator
    }
    text += own
    return true
  })
  return text
}

function noteText(
  note: ProseMirrorNode,
  blockSeparator: string | undefined,
  leafText: string | ((leaf: ProseMirrorNode) => string),
  hide: string | undefined,
): string {
  if (hide !== undefined) return hide.repeat(note.nodeSize)
  const inner = textWithoutDeletions(note, blockSeparator, leafText)
  return inner === '' ? '' : ` ${inner}`
}

function ownText(
  child: ProseMirrorNode,
  leafText: string | ((leaf: ProseMirrorNode) => string),
  hide: string | undefined,
): string {
  if (child.isText) {
    if (!isDeleted(child)) return child.text ?? ''
    return hide === undefined ? '' : hide.repeat(child.text?.length ?? 0)
  }
  if (!child.isLeaf) return ''
  return typeof leafText === 'string' ? leafText : leafText(child)
}

export const CountWithoutDeletions = CharacterCount.extend({
  onBeforeCreate(event) {
    this.parent?.(event)
    this.storage.characters = (options) => {
      const node = options?.node ?? this.editor.state.doc
      if ((options?.mode ?? this.options.mode) === 'textSize') {
        return this.options.textCounter(textWithoutDeletions(node, undefined, characterLeaf))
      }
      return node.nodeSize
    }
    this.storage.words = (options) => {
      const node = options?.node ?? this.editor.state.doc
      return this.options.wordCounter(textWithoutDeletions(node, ' ', ' ')) + equationsIn(node)
    }
  },
})

/** A equação conta uma palavra, e nenhum caractere. */
export function characterLeaf(leaf: ProseMirrorNode): string {
  return leaf.type.name === 'math' ? '' : ' '
}

/** Como no Word, uma palavra cada; no texto da contagem é um espaço. */
function equationsIn(node: ProseMirrorNode): number {
  let count = node.type.name === 'math' && !isDeleted(node) ? 1 : 0
  node.descendants((child) => {
    if (child.type.name === 'math' && !isDeleted(child)) count++
  })
  return count
}

/** O mesmo nome cai sempre na mesma cor. */
export const AUTHOR_COLORS = 6

export function authorColor(author: string | null): number {
  let hash = 0
  for (const char of author ?? '') hash = (hash * 31 + char.charCodeAt(0)) >>> 0
  return hash % AUTHOR_COLORS
}

export function revisionTitle(author: string | null, date: string | null): string {
  const when = date === null ? null : new Date(date)
  const shown = when === null || Number.isNaN(when.getTime()) ? date : when.toLocaleString()
  return [author, shown].filter((part): part is string => part !== null && part !== '').join(', ')
}

const revisionAttributes = {
  author: { default: null },
  date: { default: null },
  rid: { default: null },
  move: { default: null },
  moveName: { default: null },
}

function revisionMark(name: typeof INSERTION | typeof DELETION, tag: 'ins' | 'del') {
  return Mark.create({
    name,
    // O texto digitado na ponta não é revisão do outro autor.
    inclusive: false,
    // Quando o navegador apaga por conta própria, o ProseMirror relê o DOM: sem
    // esta regra o excluído voltaria sem marca. A colagem é limpa por `stripRevisions`.
    parseHTML: () => [
      {
        tag: `${tag}.revision`,
        getAttrs: (element) => ({
          author: element.getAttribute('data-author'),
          date: element.getAttribute('data-date'),
          rid: element.getAttribute('data-rid'),
          move: element.getAttribute('data-move'),
          moveName: element.getAttribute('data-move-name'),
        }),
      },
    ],
    addAttributes: () => ({ ...revisionAttributes }),
    renderHTML({ mark, HTMLAttributes }) {
      const author = attrString(mark.attrs['author'])
      const date = attrString(mark.attrs['date'])
      return [
        tag,
        mergeAttributes(
          {
            class: `revision revision-${tag} revision-author-${authorColor(author)}`,
            title: revisionTitle(author, date),
            'data-author': author,
            'data-date': date,
            'data-rid': attrString(mark.attrs['rid']),
            'data-move': attrString(mark.attrs['move']),
            'data-move-name': attrString(mark.attrs['moveName']),
          },
          Object.fromEntries(Object.entries(HTMLAttributes).filter(([key]) => !(key in revisionAttributes))),
        ),
        0,
      ]
    },
  })
}

export const Insertion = revisionMark(INSERTION, 'ins')
export const Deletion = revisionMark(DELETION, 'del')

export const BlockRevisions = Extension.create({
  name: 'blockRevisions',

  addGlobalAttributes() {
    const kindOf = (value: unknown): Record<string, string> => {
      const revision = blockRevisionOf(value)
      if (revision === null) return {}
      return {
        'data-revision': revision.kind,
        title: revisionTitle(attrString(revision.author), attrString(revision.date)),
        class: `revision-author-${authorColor(attrString(revision.author))}`,
      }
    }
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          markRevision: {
            default: null,
            parseHTML: () => null,
            renderHTML: (attrs) => kindOf(attrs['markRevision']),
          },
        },
      },
      {
        types: ['tableRow'],
        attributes: {
          rowRevision: {
            default: null,
            parseHTML: () => null,
            renderHTML: (attrs) => kindOf(attrs['rowRevision']),
          },
        },
      },
    ]
  },
})

export const TrackChanges = [Insertion, Deletion, BlockRevisions] as const

/** Um caractere inteiro: o par substituto de um emoji conta como um. */
export function characterSize(text: string, backward: boolean): number {
  const unit = backward ? text.charCodeAt(text.length - 1) : text.charCodeAt(0)
  const surrogate = backward ? unit >= 0xdc00 && unit <= 0xdfff : unit >= 0xd800 && unit <= 0xdbff
  return surrogate && text.length > 1 ? 2 : 1
}
