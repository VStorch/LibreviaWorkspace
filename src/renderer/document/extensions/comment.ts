import { Extension, Node } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode, type Schema } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

/**
 * Comentários (M10): as duas pontas da âncora como nós sem largura.
 *
 * Mesmo desenho do marcador (ver bookmark.ts): no arquivo a âncora é um par
 * `w:commentRangeStart`/`w:commentRangeEnd` que pode atravessar parágrafos, e é
 * por ser nó que ela sobrevive à edição do parágrafo e volta ao arquivo. O `cid`
 * é o `w:id`, que casa as pontas com o corpo em `DocumentModel.comments`.
 *
 * Uma âncora por conversa: a resposta não tem nó, e quem grava devolve as pontas
 * dela ao lado das do comentário que ela responde. O comentário de ponto — só a
 * referência, sem trecho — tem só o `commentEnd`.
 *
 * O realce do trecho é decoração da tela, e nunca vai ao JSON nem ao papel.
 */

export interface CommentAnchor {
  readonly cid: string
  /** Posição do nó de início, quando ele está no documento. */
  readonly start: number | null
  /** Posição do nó de fim, quando ele está no documento. */
  readonly end: number | null
}

/** As âncoras do documento, pelo `cid`. */
export function commentAnchorsOf(doc: ProseMirrorNode): Map<string, CommentAnchor> {
  const anchors = new Map<string, CommentAnchor>()
  doc.descendants((node, pos) => {
    const kind = node.type.name
    if (kind !== 'commentStart' && kind !== 'commentEnd') return true
    const cid = String(node.attrs['cid'] ?? '')
    const known = anchors.get(cid) ?? { cid, start: null, end: null }
    anchors.set(cid, kind === 'commentStart' ? { ...known, start: pos } : { ...known, end: pos })
    return false
  })
  return anchors
}

/**
 * As pontas de um comentário novo na seleção: o começo no `from`, o fim no `to`.
 * Sem seleção, é comentário de ponto — só o fim, como o Word grava.
 *
 * Devolve `false` quando a seleção não cai em texto (uma tabela inteira, uma
 * imagem de bloco), onde um nó de linha não cabe.
 */
export function insertCommentAnchors(tr: Transaction, schema: Schema, cid: string): boolean {
  const { from, to, $from, $to } = tr.selection
  if (!$from.parent.inlineContent || !$to.parent.inlineContent) return false
  const start = schema.nodes['commentStart']
  const end = schema.nodes['commentEnd']
  if (start === undefined || end === undefined) return false
  // O fim primeiro: inserido antes, ele empurraria a posição do começo.
  tr.insert(to, end.create({ cid }))
  if (to > from) tr.insert(from, start.create({ cid }))
  return true
}

/** Tira as pontas da conversa `cid` do texto. Devolve se havia alguma. */
export function removeCommentAnchors(tr: Transaction, cid: string): boolean {
  const positions: number[] = []
  tr.doc.descendants((node, pos) => {
    if ((node.type.name === 'commentStart' || node.type.name === 'commentEnd') && node.attrs['cid'] === cid) {
      positions.push(pos)
    }
    return node.isBlock
  })
  // De trás para a frente: cada remoção mexe só no que vem depois dela.
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
    // A âncora não é texto: não vai para a área de transferência como texto nem
    // conta palavra.
    renderText: () => '',

    addAttributes() {
      return { cid: { default: '', parseHTML: (element) => element.getAttribute('data-cid') ?? '' } }
    },

    parseHTML() {
      return [{ tag: `span[data-${name === 'commentStart' ? 'comment-start' : 'comment-end'}]` }]
    },

    renderHTML({ node }) {
      // Vazio e sem largura: na tela e no papel a âncora não ocupa lugar.
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

/** O que o painel diz ao realce: a conversa em foco e as resolvidas. */
export interface CommentFocus {
  readonly active: string | null
  readonly resolved: ReadonlySet<string>
}

interface CommentsState extends CommentFocus {
  readonly decorations: DecorationSet
}

export const commentsKey = new PluginKey<CommentsState>('comments')

/** Pede à visão que realce a conversa `cid` (ou nenhuma) — ver `Comments`. */
export function focusComment(tr: Transaction, focus: Partial<CommentFocus>): Transaction {
  return tr.setMeta(commentsKey, focus)
}

function decorate(doc: ProseMirrorNode, focus: CommentFocus): DecorationSet {
  const decorations: Decoration[] = []
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
 * O trecho colado sem âncora de comentário.
 *
 * A cópia de uma âncora seria a mesma conversa em dois lugares — o Word recusa o
 * id repetido. O que foi
 * **arrastado** dentro do documento se move, e leva a âncora junto.
 */
export function withoutCommentAnchors(slice: Slice, moving = false): Slice {
  if (moving) return slice
  let found = false
  const strip = (fragment: Fragment): Fragment => {
    const children: ProseMirrorNode[] = []
    fragment.forEach((child) => {
      if (child.type.name === 'commentStart' || child.type.name === 'commentEnd') {
        found = true
        return
      }
      children.push(child.isLeaf ? child : child.copy(strip(child.content)))
    })
    return Fragment.from(children)
  }
  const content = strip(slice.content)
  return found ? new Slice(content, slice.openStart, slice.openEnd) : slice
}

/** O realce do trecho comentado e a colagem sem âncora. */
export const Comments = Extension.create({
  name: 'comments',

  addProseMirrorPlugins() {
    return [
      new Plugin<CommentsState>({
        key: commentsKey,
        state: {
          init: (_config, state: EditorState) => {
            const focus: CommentFocus = { active: null, resolved: new Set() }
            return { ...focus, decorations: decorate(state.doc, focus) }
          },
          apply: (tr, previous) => {
            const meta = tr.getMeta(commentsKey) as Partial<CommentFocus> | undefined
            if (meta === undefined && !tr.docChanged) return previous
            const focus: CommentFocus = {
              active: meta?.active === undefined ? previous.active : meta.active,
              resolved: meta?.resolved ?? previous.resolved,
            }
            return { ...focus, decorations: decorate(tr.doc, focus) }
          },
        },
        props: {
          decorations: (state) => commentsKey.getState(state)?.decorations ?? null,
          transformPasted: (slice, view) => withoutCommentAnchors(slice, view.dragging?.move === true),
        },
      }),
    ]
  },
})
