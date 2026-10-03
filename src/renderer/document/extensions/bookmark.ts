import { Extension, Node } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { hiddenBookmarkName, nextBookmarkId } from '@services/document/bookmarks.js'
import { fieldArgument, fieldKind, fieldSwitch } from '@services/document/fields.js'

/**
 * Dois nós, e não uma marca: o par `w:bookmarkStart`/`w:bookmarkEnd` atravessa
 * parágrafos e muitas vezes cobre um trecho vazio. O `bid` é o `w:id`. Os ocultos
 * (`_Toc…`, `_Ref…`) são iguais: o sumário e as referências do Word os citam.
 */

export interface BookmarkEntry {
  readonly name: string
  readonly bid: string
  readonly pos: number
  /** Quando está no documento. */
  readonly end: number | null
}

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

/** Também os das pontas finais, cuja ponta inicial pode estar noutra parte. */
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

/** `false` quando ele não existe. */
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
 * Como o Word: o link e a referência cruzada a um título apontam para um `_Ref…`
 * em volta do texto, e o sumário para um `_Toc…`. Outro prefixo não serve:
 * "Atualizar sumário" recria os `_Toc`.
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
    // Não vai à área de transferência como texto nem conta palavra.
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
      /** O nome que já existe **muda de lugar**, como no Word. */
      setBookmark: (name: string) => ReturnType
      /** As duas pontas; o texto fica. */
      deleteBookmark: (name: string) => ReturnType
    }
  }
}

/** De trás para a frente, para as posições valerem. */
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

/** O que já tinha o nome sai antes: o marcador é um só. */
export function placeBookmark(tr: Transaction, name: string): void {
  removeBookmark(tr, name)
  const { from, to } = tr.selection
  const bid = nextBookmarkId(idsOf(tr.doc))
  const schema = tr.doc.type.schema
  // O fim primeiro: o começo deslocaria a posição dele.
  tr.insert(to, schema.nodes['bookmarkEnd']!.create({ bid }))
  tr.insert(from, schema.nodes['bookmarkStart']!.create({ name, bid }))
}

/** Dois de mesmo nome são âncora ambígua. O que foi **recortado** volta inteiro. */
export function withoutRepeatedBookmarks(slice: Slice, doc: ProseMirrorNode, moving = false): Slice {
  // Arrastar move: a origem sai na mesma transação.
  if (moving) return slice

  const names = new Set<string>()
  const ids = new Set<string>()
  doc.descendants((node) => {
    if (node.type.name === 'bookmarkStart') names.add(String(node.attrs['name']))
    if (node.type.name === 'bookmarkStart' || node.type.name === 'bookmarkEnd')
      ids.add(String(node.attrs['bid']))
    return true
  })
  if (names.size === 0 && ids.size === 0) return slice

  // O nome repetido é cópia e sai; o id repetido é de outro documento e ganha um livre.
  const dropped = new Set<string>()
  const renamed = new Map<string, string>()
  const used = new Set(ids)
  const strip = (fragment: Fragment): Fragment => {
    const children: ProseMirrorNode[] = []
    fragment.forEach((child) => {
      const kind = child.type.name
      const bid = String(child.attrs['bid'] ?? '')
      if (kind === 'bookmarkStart') {
        if (names.has(String(child.attrs['name']))) {
          dropped.add(bid)
          return
        }
        if (used.has(bid)) {
          const fresh = nextBookmarkId(used)
          used.add(fresh)
          renamed.set(bid, fresh)
          children.push(child.type.create({ ...child.attrs, bid: fresh }))
          return
        }
        used.add(bid)
      } else if (kind === 'bookmarkEnd') {
        if (dropped.has(bid)) return
        const fresh = renamed.get(bid)
        if (fresh !== undefined) {
          children.push(child.type.create({ ...child.attrs, bid: fresh }))
          return
        }
        // Repetido, fecharia o marcador daqui no lugar errado.
        if (ids.has(bid)) return
      }
      children.push(child.isLeaf ? child : child.copy(strip(child.content)))
    })
    return Fragment.from(children)
  }

  return new Slice(strip(slice.content), slice.openStart, slice.openEnd)
}

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
          /** Com `Ctrl`, como no Word, ou com clique simples no somente leitura. */
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

          /** `\h` leva ao que cita, pelo mesmo gesto do link. */
          handleClickOn(view, _pos, node, _nodePos, event) {
            if (node.type.name !== 'field') return false
            if (!event.ctrlKey && !event.metaKey && view.editable) return false
            const instr = String(node.attrs['instr'] ?? '')
            if (!['REF', 'PAGEREF'].includes(fieldKind(instr)) || fieldSwitch(instr, 'h') === null)
              return false
            const target = fieldArgument(instr)
            return target !== null && goToBookmark(view, target)
          },

          transformPasted: (slice, view) =>
            withoutRepeatedBookmarks(slice, view.state.doc, view.dragging?.move === true),
        },
      }),
    ]
  },
})
