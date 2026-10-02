import { Extension, Mark, mergeAttributes } from '@tiptap/core'
import type { Mark as ProseMirrorMark, Node as ProseMirrorNode } from '@tiptap/pm/model'
import { TextSelection, type Transaction } from '@tiptap/pm/state'
import { canJoin } from '@tiptap/pm/transform'
import { CharacterCount } from '@tiptap/extensions'

/**
 * Controle de alterações (M10, fase 1): ler, mostrar, aceitar e rejeitar.
 *
 * A revisão de texto é marca do trecho — `insertion` e `deletion` —, com o autor,
 * a data como o arquivo a escreveu e o `w:id` original (`rid`). As duas não se
 * excluem: o `w:ins` que embrulha um `w:del` é o texto que alguém inseriu e outro
 * excluiu. A movimentação (`w:moveFrom`/`w:moveTo`) aparece como exclusão e
 * inserção, e leva `move` e `moveName` para voltar como era.
 *
 * O Enter inserido ou excluído é a revisão da marca de parágrafo (`markRevision`),
 * e a linha de tabela inserida ou excluída, a da linha (`rowRevision`) — atributos
 * do bloco, como no arquivo.
 *
 * Aceitar e rejeitar são transações comuns do editor: o desfazer as devolve. O
 * controle do que se digita (fase 2) mora em track-input.ts.
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

/** A revisão de bloco, como o leitor a dá. */
export interface BlockRevision {
  readonly kind: 'ins' | 'del'
  readonly author?: string
  readonly date?: string
  readonly rid?: string
}

export type ChangeKind =
  'insertion' | 'deletion' | 'markInsertion' | 'markDeletion' | 'rowInsertion' | 'rowDeletion'

/** Uma alteração: o trecho de texto, a marca de parágrafo ou a linha. */
export interface RevisionChange {
  readonly kind: ChangeKind
  readonly from: number
  readonly to: number
  readonly author: string | null
  readonly date: string | null
  /** Os pedaços de texto do trecho, sem as pontas sem largura entre eles. */
  readonly segments: readonly { readonly from: number; readonly to: number }[]
  /** A marca do trecho — só nas alterações de texto. */
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

/** O mesmo autor e o mesmo `rid` (o arquivo), ou o mesmo autor sem `rid` (o editor). */
function sameRevision(a: ProseMirrorMark, b: ProseMirrorMark): boolean {
  return a.attrs['author'] === b.attrs['author'] && a.attrs['rid'] === b.attrs['rid']
}

/**
 * Todas as alterações do documento, na ordem do texto.
 *
 * Um trecho é o texto contíguo com a mesma marca (mesmo autor e mesmo `rid`); as
 * pontas de marcador e de comentário no meio não o partem — e sobrevivem ao
 * aceite, porque só os pedaços de texto saem.
 */
export function revisionChangesOf(doc: ProseMirrorNode): RevisionChange[] {
  const changes: RevisionChange[] = []
  collectChanges(doc, 0, changes)
  return changes.sort((a, b) => a.from - b.from || a.to - b.to)
}

/**
 * As alterações dentro de `parent`, cujo conteúdo começa em `base`. Desce nos
 * corpos de nota (M11): o que se controla numa nota é alteração do documento,
 * e aceitar ou rejeitar todas não pode deixá-la para trás.
 */
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

/**
 * A alteração no cursor: o trecho que o contém, senão a marca do parágrafo dele,
 * senão a linha. `null` fora de qualquer alteração.
 */
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

/** O que aceitar faz com cada alteração — e rejeitar faz o contrário. */
type Effect = 'keep' | 'drop'

function effectOf(kind: ChangeKind, accept: boolean): Effect {
  const inserted = kind === INSERTION || kind === 'markInsertion' || kind === 'rowInsertion'
  return inserted === accept ? 'keep' : 'drop'
}

/**
 * Aplica o aceite ou a rejeição de uma alteração à transação. As posições da
 * alteração são as do documento em que ela foi achada, e passam pelo mapeamento
 * dos passos que a transação já tem desde `since`.
 */
function settle(tr: Transaction, change: RevisionChange, accept: boolean, since: number): void {
  const mapping = tr.mapping.slice(since)
  const effect = effectOf(change.kind, accept)

  switch (change.kind) {
    case INSERTION:
    case DELETION: {
      const segments = change.segments
        .map((segment) => ({ from: mapping.map(segment.from, 1), to: mapping.map(segment.to, -1) }))
        .filter((segment) => segment.to > segment.from)
      // De trás para frente: apagar um pedaço não mexe nas posições dos anteriores.
      for (const segment of [...segments].reverse()) {
        if (effect === 'drop') tr.delete(segment.from, segment.to)
        else tr.removeMark(segment.from, segment.to, change.mark)
      }
      return
    }

    case 'markInsertion':
    case 'markDeletion': {
      const end = mapping.map(change.from)
      const paragraph = tr.doc.resolve(end).parent
      const start = end - paragraph.nodeSize + 1
      if (start < 0 || tr.doc.nodeAt(start) !== paragraph) return
      tr.setNodeMarkup(start, undefined, { ...paragraph.attrs, markRevision: null })
      // O Enter que sai junta este parágrafo ao seguinte, como no Word.
      const after = start + paragraph.nodeSize
      if (effect === 'drop' && after < tr.doc.content.size && canJoin(tr.doc, after)) tr.join(after)
      return
    }

    case 'rowInsertion':
    case 'rowDeletion': {
      const from = mapping.map(change.from, 1)
      const row = tr.doc.nodeAt(from)
      if (row === null || row.type.name !== 'tableRow') return
      if (effect === 'keep') {
        tr.setNodeMarkup(from, undefined, { ...row.attrs, rowRevision: null })
        return
      }
      const $row = tr.doc.resolve(from)
      const table = $row.parent
      // A última linha leva a tabela junto: tabela sem linha não existe.
      if (table.childCount === 1) {
        const tableStart = $row.before()
        tr.delete(tableStart, tableStart + table.nodeSize)
      } else {
        tr.delete(from, from + row.nodeSize)
      }
      return
    }
  }
}

/** Aceita (ou rejeita) a alteração no cursor. Devolve se havia alguma. */
export function settleChangeAt(tr: Transaction, pos: number, accept: boolean): boolean {
  const change = changeAt(tr.doc, pos)
  if (change === null) return false
  settle(tr, change, accept, tr.steps.length)
  return true
}

/** Aceita (ou rejeita) todas, numa transação só — um desfazer devolve tudo. */
export function settleAllChanges(tr: Transaction, accept: boolean): boolean {
  const changes = revisionChangesOf(tr.doc)
  if (changes.length === 0) return false
  const since = tr.steps.length
  // Do fim para o começo: o que se apaga adiante não mexe no que vem antes.
  for (const change of [...changes].reverse()) settle(tr, change, accept, since)
  return true
}

/** A alteração seguinte (ou a anterior) ao cursor, pela ordem do texto. */
export function adjacentChange(doc: ProseMirrorNode, pos: number, direction: 1 | -1): RevisionChange | null {
  const changes = revisionChangesOf(doc)
  // Para frente, a que começa no cursor também vale — a menos que acabe nele, que
  // é a que acabou de ser escolhida. Para trás, a que acaba antes do cursor: a
  // linha escolhida põe o cursor dentro dela, e "começa antes" a escolheria de novo.
  if (direction === 1) {
    return changes.find((change) => change.from > pos || (change.from === pos && change.to > pos)) ?? null
  }
  return [...changes].reverse().find((change) => change.from < pos && change.to <= pos) ?? null
}

/** Escolhe a alteração na tela: o trecho selecionado, ou o cursor na marca ou na linha. */
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

// --- texto sem o excluído --------------------------------------------------

function isDeleted(node: ProseMirrorNode): boolean {
  return node.marks.some((mark) => mark.type.name === DELETION)
}

/**
 * O texto do nó, como `textBetween`, mas sem o que está marcado como excluído.
 *
 * Com `hide`, o excluído vira esse caractere repetido em vez de sumir: a busca
 * precisa do texto com o comprimento das posições, e um caractere que nunca casa
 * impede que ela ache um termo atravessando o trecho excluído.
 */
export function textWithoutDeletions(
  node: ProseMirrorNode,
  blockSeparator: string | undefined,
  leafText: string,
  hide?: string,
): string {
  let text = ''
  let first = true
  node.descendants((child) => {
    // A nota (M11) é um nó só no texto do parágrafo: a busca precisa do
    // comprimento dela nas posições, e o corpo não está na tela para ser achado.
    // Na contagem ela entra — o Word conta as notas —, separada do texto em volta.
    if (child.type.name === 'noteRef') {
      if (hide !== undefined) text += hide.repeat(child.nodeSize)
      else {
        const inner = textWithoutDeletions(child, blockSeparator, leafText)
        if (inner !== '') text += ` ${inner}`
      }
      return false
    }
    const own = child.isText
      ? isDeleted(child)
        ? hide === undefined
          ? ''
          : hide.repeat(child.text?.length ?? 0)
        : (child.text ?? '')
      : child.isLeaf
        ? leafText
        : ''
    if (
      child.isBlock &&
      ((child.isLeaf && own !== '') || child.isTextblock) &&
      blockSeparator !== undefined
    ) {
      if (first) first = false
      else text += blockSeparator
    }
    text += own
    return true
  })
  return text
}

/** A contagem de palavras e caracteres, sem o texto excluído. Mesmo nome e mesmo armazenamento. */
export const CountWithoutDeletions = CharacterCount.extend({
  onBeforeCreate(event) {
    this.parent?.(event)
    this.storage.characters = (options) => {
      const node = options?.node ?? this.editor.state.doc
      if ((options?.mode ?? this.options.mode) === 'textSize') {
        return this.options.textCounter(textWithoutDeletions(node, undefined, ' '))
      }
      return node.nodeSize
    }
    this.storage.words = (options) => {
      const node = options?.node ?? this.editor.state.doc
      return this.options.wordCounter(textWithoutDeletions(node, ' ', ' ')) + equationsIn(node)
    }
  },
})

/**
 * As equações do trecho (M11), que contam uma palavra cada, como no Word. No texto
 * da contagem ela é um espaço — o LaTeX de uma ou o marcador "[equação]" inflariam
 * a conta de palavras com o que ninguém escreveu.
 */
function equationsIn(node: ProseMirrorNode): number {
  let count = node.type.name === 'math' && !isDeleted(node) ? 1 : 0
  node.descendants((child) => {
    if (child.type.name === 'math' && !isDeleted(child)) count++
  })
  return count
}

// --- as marcas e os atributos ----------------------------------------------

/** Uma cor por autor, estável entre aberturas: o mesmo nome cai sempre na mesma. */
export const AUTHOR_COLORS = 6

export function authorColor(author: string | null): number {
  let hash = 0
  for (const char of author ?? '') hash = (hash * 31 + char.charCodeAt(0)) >>> 0
  return hash % AUTHOR_COLORS
}

/** "autor, data" — a dica que o trecho mostra sob o ponteiro. */
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
    // A revisão não se estende ao que se digita na ponta dela: texto novo não é
    // revisão do outro autor (o controle do que se digita põe a marca à mão — ver track-input.ts).
    inclusive: false,
    // A tela relida: quando o navegador apaga ou digita por conta própria (o
    // Backspace num caractere, o Ctrl+Backspace), o ProseMirror relê o trecho do
    // DOM — sem esta regra o excluído voltava sem marca e era tomado por texto
    // novo. A colagem não traz revisão: `stripRevisions` (track-input.ts) a tira.
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
            // Para a releitura da tela (ver `parseHTML`).
            'data-author': author,
            'data-date': date,
            'data-rid': attrString(mark.attrs['rid']),
            'data-move': attrString(mark.attrs['move']),
            'data-move-name': attrString(mark.attrs['moveName']),
          },
          // Os atributos crus ficam fora do HTML.
          Object.fromEntries(Object.entries(HTMLAttributes).filter(([key]) => !(key in revisionAttributes))),
        ),
        0,
      ]
    },
  })
}

export const Insertion = revisionMark(INSERTION, 'ins')
export const Deletion = revisionMark(DELETION, 'del')

/** Os atributos de bloco: a marca de parágrafo e a linha de tabela revisadas. */
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
