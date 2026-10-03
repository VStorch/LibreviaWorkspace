import { Extension, Node } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode, type Schema } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { noteRefAround } from './note-ref.js'
import { KEEP_SELECTION } from './zero-width.js'

/**
 * Como o marcador (`bookmark.ts`): a âncora `w:commentRangeStart`/`End` atravessa
 * parágrafos, e é por ser nó que volta ao arquivo. O `cid` é o `w:id`. Uma âncora
 * por conversa: a resposta não tem nó. O comentário de ponto tem só o
 * `commentEnd`. O realce é decoração, e nunca vai ao JSON nem ao papel.
 */

export interface CommentAnchor {
  readonly cid: string
  /** Quando está no documento. */
  readonly start: number | null
  readonly end: number | null
}

export function commentAnchorsOf(doc: ProseMirrorNode): Map<string, CommentAnchor> {
  const anchors = new Map<string, CommentAnchor>()
  doc.descendants((node, pos) => {
    const kind = node.type.name
    // Desce no corpo das notas: a conversa numa nota também tem cartão.
    if (kind !== 'commentStart' && kind !== 'commentEnd') return true
    const cid = String(node.attrs['cid'] ?? '')
    const known = anchors.get(cid) ?? { cid, start: null, end: null }
    anchors.set(cid, kind === 'commentStart' ? { ...known, start: pos } : { ...known, end: pos })
    return false
  })
  return anchors
}

/** O trecho entre as pontas; no comentário de ponto, o cursor na ponta. */
export function commentSelectionOf(anchor: CommentAnchor): { readonly from: number; readonly to: number } {
  const from = anchor.start === null ? (anchor.end ?? 0) : anchor.start + 1
  const to = anchor.end ?? from
  return { from, to: Math.max(from, to) }
}

/**
 * `threads` são as conversas com cartão. A partir da conversa em foco, se o
 * cursor está nela; senão, do cursor; circular. `null` sem conversa no texto.
 */
export function adjacentComment(
  doc: ProseMirrorNode,
  threads: ReadonlySet<string>,
  active: string | null,
  cursor: number,
  direction: 1 | -1,
): string | null {
  const order = [...commentAnchorsOf(doc).values()]
    .filter((anchor) => threads.has(anchor.cid))
    .map((anchor) => ({
      cid: anchor.cid,
      pos: (anchor.start ?? anchor.end)!,
      end: (anchor.end ?? anchor.start)!,
    }))
    .sort((left, right) => left.pos - right.pos)
  if (order.length === 0) return null
  // A conversa em foco vale enquanto o cursor está no trecho dela.
  const current = order.findIndex(
    (item) => item.cid === active && cursor >= item.pos && cursor <= item.end + 1,
  )
  if (current >= 0) return order[(current + direction + order.length) % order.length]!.cid
  const found =
    direction > 0
      ? order.find((item) => item.pos >= cursor)
      : [...order].reverse().find((item) => item.pos < cursor)
  return (found ?? (direction > 0 ? order[0] : order[order.length - 1]))!.cid
}

/**
 * Sem seleção, comentário de ponto, só o fim, como o Word grava. `range` vem de
 * uma nota (`caretOf`). `false` quando a seleção não cai em texto.
 */
export function insertCommentAnchors(
  tr: Transaction,
  schema: Schema,
  cid: string,
  range: { readonly from: number; readonly to: number } = tr.selection,
): boolean {
  const { from, to } = range
  const $from = tr.doc.resolve(from)
  const $to = tr.doc.resolve(to)
  if (!$from.parent.inlineContent || !$to.parent.inlineContent) return false
  const start = schema.nodes['commentStart']
  const end = schema.nodes['commentEnd']
  if (start === undefined || end === undefined) return false
  // O fim primeiro: ele empurraria a posição do começo.
  tr.insert(to, end.create({ cid }))
  if (to > from) tr.insert(from, start.create({ cid }))
  return true
}

/** Devolve se havia alguma. */
export function removeCommentAnchors(tr: Transaction, cid: string): boolean {
  const positions: number[] = []
  tr.doc.descendants((node, pos) => {
    if ((node.type.name === 'commentStart' || node.type.name === 'commentEnd') && node.attrs['cid'] === cid) {
      positions.push(pos)
    }
    // O corpo da nota é filho de um nó em linha.
    return node.isBlock || node.type.name === 'noteRef'
  })
  for (const pos of positions.reverse()) tr.delete(pos, pos + 1)
  return positions.length > 0
}

const commentNode = (name: 'commentStart' | 'commentEnd') =>
  Node.create({
    name,
    group: 'inline',
    inline: true,
    atom: true,
    selectable: false,
    // Não vai à área de transferência como texto nem conta palavra.
    renderText: () => '',

    addAttributes() {
      return { cid: { default: '', parseHTML: (element) => element.getAttribute('data-cid') ?? '' } }
    },

    parseHTML() {
      return [{ tag: `span[data-${name === 'commentStart' ? 'comment-start' : 'comment-end'}]` }]
    },

    renderHTML({ node }) {
      return [
        'span',
        {
          [`data-${name === 'commentStart' ? 'comment-start' : 'comment-end'}`]: '',
          'data-cid': String(node.attrs['cid']),
          class: 'comment-anchor',
        },
      ]
    },
  })

export const CommentStart = commentNode('commentStart')
export const CommentEnd = commentNode('commentEnd')

export interface CommentFocus {
  readonly active: string | null
  readonly resolved: ReadonlySet<string>
  readonly hidden: boolean
}

interface CommentsState extends CommentFocus {
  readonly decorations: DecorationSet
}

export const commentsKey = new PluginKey<CommentsState>('comments')

export function focusComment(tr: Transaction, focus: Partial<CommentFocus>): Transaction {
  return tr.setMeta(commentsKey, focus)
}

/** O realce em foco e o trecho selecionado, à vista. */
export function selectComment(tr: Transaction, cid: string): Transaction {
  const next = focusComment(tr, { active: cid })
  const anchor = commentAnchorsOf(tr.doc).get(cid)
  if (anchor === undefined) return next
  const { from, to } = commentSelectionOf(anchor)
  // Numa nota o trecho é escolhido no corpo (`selectInNote`), que tem editor próprio.
  if (noteRefAround(tr.doc, from) !== null) return next
  return next
    .setSelection(TextSelection.create(tr.doc, from, to))
    .setMeta(KEEP_SELECTION, true)
    .scrollIntoView()
}

function decorate(doc: ProseMirrorNode, focus: CommentFocus): DecorationSet {
  const decorations: Decoration[] = []
  if (focus.hidden) return DecorationSet.empty
  for (const anchor of commentAnchorsOf(doc).values()) {
    if (anchor.start === null || anchor.end === null || anchor.end <= anchor.start + 1) continue
    const classes = ['comment-range']
    if (focus.resolved.has(anchor.cid)) classes.push('comment-range--resolved')
    if (focus.active === anchor.cid) classes.push('comment-range--active')
    decorations.push(
      Decoration.inline(anchor.start + 1, anchor.end, { class: classes.join(' '), 'data-cid': anchor.cid }),
    )
  }
  return DecorationSet.create(doc, decorations)
}

/**
 * A cópia de uma âncora seria a mesma conversa em dois lugares: sai. O que foi
 * **recortado** volta inteiro, ponta a ponta; o arrastado se move. `isKnown` tira
 * também a âncora de outro documento, sem corpo aqui.
 */
export function withoutCommentAnchors(
  slice: Slice,
  doc: ProseMirrorNode,
  moving = false,
  isKnown: (cid: string) => boolean = () => true,
): Slice {
  if (moving) return slice
  const present = new Set<string>()
  doc.descendants((node) => {
    if (node.type.name === 'commentStart' || node.type.name === 'commentEnd')
      present.add(`${node.type.name}:${String(node.attrs['cid'])}`)
    return true
  })
  let dropped = false
  const strip = (fragment: Fragment): Fragment => {
    const children: ProseMirrorNode[] = []
    fragment.forEach((child) => {
      const kind = child.type.name
      if (kind === 'commentStart' || kind === 'commentEnd') {
        const cid = String(child.attrs['cid'] ?? '')
        if (present.has(`${kind}:${cid}`) || !isKnown(cid)) {
          dropped = true
          return
        }
      }
      children.push(child.isLeaf ? child : child.copy(strip(child.content)))
    })
    return Fragment.from(children)
  }
  const content = strip(slice.content)
  return dropped ? new Slice(content, slice.openStart, slice.openEnd) : slice
}

export interface CommentsOptions {
  readonly isKnown: ((cid: string) => boolean) | undefined
}

export const Comments = Extension.create<CommentsOptions>({
  name: 'comments',

  addOptions() {
    return { isKnown: undefined }
  },

  addProseMirrorPlugins() {
    const isKnown = this.options.isKnown
    return [
      new Plugin<CommentsState>({
        key: commentsKey,
        state: {
          init: (_config, state: EditorState) => {
            const focus: CommentFocus = { active: null, resolved: new Set(), hidden: false }
            return { ...focus, decorations: decorate(state.doc, focus) }
          },
          apply: (tr, previous) => {
            const meta = tr.getMeta(commentsKey) as Partial<CommentFocus> | undefined
            if (meta === undefined && !tr.docChanged) return previous
            const focus: CommentFocus = {
              active: meta?.active === undefined ? previous.active : meta.active,
              resolved: meta?.resolved ?? previous.resolved,
              hidden: meta?.hidden ?? previous.hidden,
            }
            return { ...focus, decorations: decorate(tr.doc, focus) }
          },
        },
        props: {
          decorations: (state) => commentsKey.getState(state)?.decorations ?? null,
          transformPasted: (slice, view) =>
            withoutCommentAnchors(slice, view.state.doc, view.dragging?.move === true, isKnown),
        },
      }),
    ]
  },
})
