import { Extension, Node } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { hiddenBookmarkName, nextBookmarkId } from '@services/document/bookmarks.js'

/**
 * Marcadores (bookmarks): as duas pontas como nós sem largura.
 *
 * No arquivo o marcador é um par `w:bookmarkStart`/`w:bookmarkEnd` que pode
 * começar num parágrafo e terminar noutro — por isso dois nós, e não uma marca de
 * texto: a marca não atravessa parágrafo e não existe sobre um trecho vazio, que
 * é justamente o marcador mais comum (o ponto onde o cursor estava). O `bid` é o
 * `w:id` do arquivo, que casa as pontas; o nome mora só no começo, como lá.
 *
 * Os ocultos (`_Toc…`, `_Ref…`, `_GoBack`) são nós iguais aos outros: o sumário e
 * as referências cruzadas do Word apontam para eles, e é por isso que precisam
 * sobreviver à edição do parágrafo que os carrega.
 */

export interface BookmarkEntry {
  readonly name: string
  readonly bid: string
  /** Posição do nó de início. */
  readonly pos: number
  /** Posição do nó de fim, quando ele está no documento. */
  readonly end: number | null
}

/** Os marcadores do documento, na ordem em que começam. */
export function bookmarksOf(doc: ProseMirrorNode): BookmarkEntry[] {
  const starts: Array<{ name: string; bid: string; pos: number }> = []
  const ends = new Map<string, number>()

  doc.descendants((node, pos) => {
    if (node.type.name === 'bookmarkStart') {
      starts.push({ name: String(node.attrs['name'] ?? ''), bid: String(node.attrs['bid'] ?? ''), pos })
    } else if (node.type.name === 'bookmarkEnd') {
      ends.set(String(node.attrs['bid'] ?? ''), pos)
    }
    return true
  })

  return starts.map((start) => ({ ...start, end: ends.get(start.bid) ?? null }))
}

/** Os ids em uso — os das pontas finais também, cuja ponta inicial pode estar noutra parte. */
function idsOf(doc: ProseMirrorNode): string[] {
  const ids: string[] = []
  doc.descendants((node) => {
    if (node.type.name === 'bookmarkStart' || node.type.name === 'bookmarkEnd') {
      ids.push(String(node.attrs['bid'] ?? ''))
    }
    return true
  })
  return ids
}

/** Leva o cursor ao marcador e rola a folha até ele. `false` quando ele não existe. */
export function goToBookmark(view: EditorView, name: string): boolean {
  const entry = bookmarksOf(view.state.doc).find((bookmark) => bookmark.name === name)
  if (entry === undefined) return false

  const to = entry.end !== null && entry.end > entry.pos ? entry.end : entry.pos + 1
  view.dispatch(
    view.state.tr.setSelection(
      TextSelection.create(view.state.doc, entry.pos + 1, Math.max(to, entry.pos + 1)),
    ),
  )
  view.focus()
  const dom = view.nodeDOM(entry.pos)
  if (dom instanceof HTMLElement) dom.scrollIntoView({ block: 'center' })
  return true
}

/**
 * O nome do marcador que envolve o texto do bloco em `pos` — e, se não houver, um
 * oculto novo com o prefixo dado, que é como o Word faz: o link para um título e a
 * referência cruzada a ele apontam para um `_Ref…` em volta do texto, e o sumário
 * para um `_Toc…`.
 *
 * Aproveita o que já está no começo do bloco com o mesmo prefixo; outro prefixo
 * não serve, porque "Atualizar sumário" recria os `_Toc` e levaria junto a
 * referência que apontasse para eles.
 */
export function ensureBlockBookmark(view: EditorView, pos: number, prefix: '_Ref' | '_Toc'): string | null {
  const block = view.state.doc.nodeAt(pos)
  if (block === null || !block.isTextblock) return null

  let found: string | null = null
  block.forEach((child) => {
    const name = String(child.attrs['name'] ?? '')
    if (found === null && child.type.name === 'bookmarkStart' && name.startsWith(prefix)) found = name
  })
  if (found !== null) return found

  const existing = bookmarksOf(view.state.doc)
  const name = hiddenBookmarkName(
    prefix,
    existing.map((bookmark) => bookmark.name),
  )
  const bid = nextBookmarkId(idsOf(view.state.doc))
  const schema = view.state.schema
  const tr = view.state.tr
  tr.insert(pos + block.nodeSize - 1, schema.nodes['bookmarkEnd']!.create({ bid }))
  tr.insert(pos + 1, schema.nodes['bookmarkStart']!.create({ name, bid }))
  view.dispatch(tr)
  return name
}

const bookmarkNode = (name: 'bookmarkStart' | 'bookmarkEnd') =>
  Node.create({
    name,
    group: 'inline',
    inline: true,
    atom: true,
    selectable: false,
    // O marcador não é texto: não vai para a área de transferência como texto
    // nem conta palavra.
    renderText: () => '',

    addAttributes() {
      return name === 'bookmarkStart'
        ? {
            name: { default: '', parseHTML: (element) => element.getAttribute('data-bookmark') ?? '' },
            bid: { default: '', parseHTML: (element) => element.getAttribute('data-bid') ?? '' },
          }
        : { bid: { default: '', parseHTML: (element) => element.getAttribute('data-bid') ?? '' } }
    },

    parseHTML() {
      return [{ tag: `span[data-${name === 'bookmarkStart' ? 'bookmark' : 'bookmark-end'}]` }]
    },

    renderHTML({ node }) {
      // Vazio e sem largura: na tela e no papel o marcador não ocupa lugar.
      return name === 'bookmarkStart'
        ? [
            'span',
            {
              'data-bookmark': String(node.attrs['name']),
              'data-bid': String(node.attrs['bid']),
              class: 'bookmark',
            },
          ]
        : ['span', { 'data-bookmark-end': '', 'data-bid': String(node.attrs['bid']), class: 'bookmark' }]
    },
  })

export const BookmarkStart = bookmarkNode('bookmarkStart')
export const BookmarkEnd = bookmarkNode('bookmarkEnd')

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    bookmarks: {
      /**
       * Marca a seleção com o nome dado. O nome que já existe **muda de lugar**,
       * como no Word: o marcador é um só.
       */
      setBookmark: (name: string) => ReturnType
      /** Tira o marcador do documento — as duas pontas. O texto fica. */
      deleteBookmark: (name: string) => ReturnType
    }
  }
}

/** Remove as pontas de um marcador, de trás para a frente para as posições valerem. */
export function removeBookmark(tr: Transaction, name: string): boolean {
  const entry = bookmarksOf(tr.doc).find((bookmark) => bookmark.name === name)
  if (entry === undefined) return false

  const positions: number[] = []
  tr.doc.descendants((node, pos) => {
    const kind = node.type.name
    if ((kind === 'bookmarkStart' || kind === 'bookmarkEnd') && String(node.attrs['bid']) === entry.bid) {
      positions.push(pos)
    }
    return true
  })
  for (const pos of positions.sort((left, right) => right - left)) tr.delete(pos, pos + 1)
  return true
}

/**
 * Marca a seleção da transação com o nome dado. O que já tinha o nome sai antes —
 * o marcador é um só —, e o id é um a mais que o maior do documento.
 */
export function placeBookmark(tr: Transaction, name: string): void {
  removeBookmark(tr, name)
  const { from, to } = tr.selection
  const bid = nextBookmarkId(idsOf(tr.doc))
  const schema = tr.doc.type.schema
  // O fim primeiro: inserir o começo antes deslocaria a posição dele.
  tr.insert(to, schema.nodes['bookmarkEnd']!.create({ bid }))
  tr.insert(from, schema.nodes['bookmarkStart']!.create({ name, bid }))
}

/**
 * O trecho colado sem os marcadores que o documento já tem.
 *
 * Copiar um parágrafo leva os marcadores dele junto, e dois de mesmo nome são
 * âncora ambígua — o Word recusa o id repetido. O que foi **recortado** não está
 * mais no documento, e volta inteiro: mover um parágrafo não custa o marcador.
 */
export function withoutRepeatedBookmarks(slice: Slice, doc: ProseMirrorNode): Slice {
  const present = new Set<string>()
  doc.descendants((node) => {
    if (node.type.name === 'bookmarkStart') present.add(`n:${String(node.attrs['name'])}`)
    if (node.type.name === 'bookmarkStart' || node.type.name === 'bookmarkEnd') {
      present.add(`i:${String(node.attrs['bid'])}`)
    }
    return true
  })
  if (present.size === 0) return slice

  const dropped = new Set<string>()
  const strip = (fragment: Fragment): Fragment => {
    const children: ProseMirrorNode[] = []
    fragment.forEach((child) => {
      const kind = child.type.name
      const bid = String(child.attrs['bid'] ?? '')
      if (
        kind === 'bookmarkStart' &&
        (present.has(`n:${String(child.attrs['name'])}`) || present.has(`i:${bid}`))
      ) {
        dropped.add(bid)
        return
      }
      if (kind === 'bookmarkEnd' && (dropped.has(bid) || present.has(`i:${bid}`))) return
      children.push(child.isLeaf ? child : child.copy(strip(child.content)))
    })
    return Fragment.from(children)
  }

  return new Slice(strip(slice.content), slice.openStart, slice.openEnd)
}

/**
 * Os comandos, o clique no link interno e a colagem sem marcador repetido.
 */
export const Bookmarks = Extension.create({
  name: 'bookmarks',

  addCommands() {
    return {
      setBookmark:
        (name: string) =>
        ({ tr, dispatch }) => {
          if (dispatch !== undefined) placeBookmark(tr, name)
          return true
        },

      deleteBookmark:
        (name: string) =>
        ({ tr, dispatch }) =>
          dispatch === undefined
            ? bookmarksOf(tr.doc).some((bookmark) => bookmark.name === name)
            : removeBookmark(tr, name),
    }
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('bookmarks'),
        props: {
          /**
           * O link para um lugar do documento leva a ele: com `Ctrl`, como no
           * Word, ou com clique simples no somente leitura, onde clicar não tem
           * outro uso. O link externo continua como era — quem o abre é o
           * processo main, depois da allowlist.
           */
          handleClick(view, pos, event) {
            if (!event.ctrlKey && !event.metaKey && view.editable) return false
            const $pos = view.state.doc.resolve(pos)
            const link = [...$pos.marks(), ...($pos.nodeAfter?.marks ?? [])].find(
              (mark) => mark.type.name === 'link',
            )
            const href = link?.attrs['href']
            if (typeof href !== 'string' || !href.startsWith('#')) return false
            event.preventDefault()
            return goToBookmark(view, href.slice(1))
          },

          // Ver `withoutRepeatedBookmarks`.
          transformPasted: (slice, view) => withoutRepeatedBookmarks(slice, view.state.doc),
        },
      }),
    ]
  },
})
