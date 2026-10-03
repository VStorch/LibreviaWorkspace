import { Node } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode, type Schema } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { DocumentNotes } from '@services/document/model.js'
import { NoteKind, noteLabels } from '@services/document/notes.js'
import { sectionBreakIn, type SectionBlock } from '@services/document/sections.js'
import { noteRefView } from './note-view.js'

/**
 * Um nó em linha e atômico, **com conteúdo**: o corpo tem editor próprio
 * (`note-view.ts`), mas viaja dentro da referência — copiar copia a nota, apagar
 * apaga. O `nid` é o `w:id`; a colada ao lado da original o perde, e ganha nota
 * própria. O número é a ordem da referência, desenhado por decoração.
 */

export interface NoteRefOptions {
  /** Consultada a cada conta. */
  readonly notes: (() => DocumentNotes | undefined) | undefined
}

export const noteRefKey = new PluginKey<NoteRefState>('noteRef')

/** Sem descer no corpo de nenhuma. */
export function noteRefsOf(doc: ProseMirrorNode): Array<{ node: ProseMirrorNode; pos: number }> {
  const found: Array<{ node: ProseMirrorNode; pos: number }> = []
  doc.descendants((node, pos) => {
    if (node.type.name !== 'noteRef') return true
    found.push({ node, pos })
    return false
  })
  return found
}

/** `null` fora de nota: é por ela que um comando sabe que a seleção é a do corpo. */
export function noteRefAround(doc: ProseMirrorNode, pos: number): number | null {
  if (pos < 0 || pos > doc.content.size) return null
  const $pos = doc.resolve(pos)
  for (let depth = $pos.depth; depth > 0; depth--) {
    if ($pos.node(depth).type.name === 'noteRef') return $pos.before(depth)
  }
  return null
}

function markOf(node: ProseMirrorNode): string | null {
  const mark: unknown = node.attrs['mark']
  return typeof mark === 'string' && mark !== '' ? mark : null
}

/** A de marca própria não: o Word grava a marca no corpo, e desenhá-la de novo daria "**". */
export function drawsNoteNumber(node: ProseMirrorNode): boolean {
  return markOf(node) === null
}

/** Pela contagem das marcas de seção antes dela, que fecham a seção como o `w:sectPr`. */
function noteRefSections(doc: ProseMirrorNode): number[] {
  const sections: number[] = []
  let section = 0
  doc.forEach((block) => {
    block.descendants((node) => {
      if (node.type.name !== 'noteRef') return true
      sections.push(section)
      return false
    })
    if (sectionBreakIn(block as unknown as SectionBlock) !== null) section += 1
  })
  return sections
}

/** `pages` é a folha de cada uma, para o reinício por página. */
export function noteRefLabels(
  doc: ProseMirrorNode,
  notes?: DocumentNotes,
  pages: readonly (number | undefined)[] = [],
): string[] {
  const sections = noteRefSections(doc)
  return noteLabels(
    noteRefsOf(doc).map(({ node }, index) => {
      const page = pages[index]
      return {
        kind: String(node.attrs['kind']),
        mark: markOf(node),
        section: sections[index] ?? 0,
        ...(page === undefined ? {} : { page }),
      }
    }),
    notes,
  )
}

interface NoteRefState {
  readonly decorations: DecorationSet
  readonly labels: readonly string[]
  readonly pages: readonly (number | undefined)[]
}

function noteRefState(
  doc: ProseMirrorNode,
  notes: DocumentNotes | undefined,
  pages: readonly (number | undefined)[],
): NoteRefState {
  const refs = noteRefsOf(doc)
  if (refs.length === 0) return { decorations: DecorationSet.empty, labels: [], pages }
  const labels = noteRefLabels(doc, notes, pages)
  const decorations = DecorationSet.create(
    doc,
    refs.map(({ node, pos }, index) =>
      Decoration.node(
        pos,
        pos + node.nodeSize,
        markOf(node) === null ? { 'data-note-number': labels[index]! } : {},
        { noteLabel: labels[index]! },
      ),
    ),
  )
  return { decorations, labels, pages }
}

/** O papel e o `NOTEREF` usam os mesmos. */
export function noteLabelsOf(state: EditorState): readonly string[] {
  // Sem o plugin (um estado montado à parte), a conta padrão do documento.
  return noteRefKey.getState(state)?.labels ?? noteRefLabels(state.doc)
}

export function notePagesOf(state: EditorState): readonly (number | undefined)[] {
  return noteRefKey.getState(state)?.pages ?? []
}

/** Vem da paginação, depois de ela assentar; a transação não muda o documento. */
export function setNotePages(tr: Transaction, pages: readonly (number | undefined)[]): Transaction {
  return tr.setMeta(noteRefKey, pages).setMeta('addToHistory', false)
}

export function footnotePagesOf(
  areas: ReadonlyArray<{
    readonly sheet: number
    readonly kind: string
    readonly items: ReadonlyArray<{ readonly index: number; readonly fromLine: number }>
  }>,
): (number | undefined)[] {
  const pages: (number | undefined)[] = []
  for (const area of areas) {
    if (area.kind !== NoteKind.Footnote) continue
    for (const item of area.items) if (item.fromLine === 0) pages[item.index] = area.sheet
  }
  return pages
}

export function samePages(
  left: readonly (number | undefined)[],
  right: readonly (number | undefined)[],
): boolean {
  const length = Math.max(left.length, right.length)
  for (let index = 0; index < length; index++) if (left[index] !== right[index]) return false
  return true
}

/**
 * O arquivo não aceita duas referências à mesma nota: a colada vai sem `nid` e
 * ganha nota própria. Arrastar não é colar, e leva o `nid` junto.
 */
export function withoutRepeatedNotes(slice: Slice, doc: ProseMirrorNode, moving = false): Slice {
  if (moving) return slice
  const present = new Set(
    noteRefsOf(doc).map(({ node }) => `${String(node.attrs['kind'])}:${String(node.attrs['nid'])}`),
  )
  let changed = false
  const strip = (fragment: Fragment): Fragment => {
    const children: ProseMirrorNode[] = []
    fragment.forEach((child) => {
      if (child.type.name === 'noteRef' && child.attrs['nid'] !== null) {
        if (present.has(`${String(child.attrs['kind'])}:${String(child.attrs['nid'])}`)) {
          changed = true
          children.push(child.type.create({ ...child.attrs, nid: null }, child.content, child.marks))
          return
        }
      }
      children.push(child.isLeaf ? child : child.copy(strip(child.content)))
    })
    return Fragment.from(children)
  }
  const content = strip(slice.content)
  return changed ? new Slice(content, slice.openStart, slice.openEnd) : slice
}

function contentFromJson(element: HTMLElement, schema: Schema): Fragment {
  try {
    const raw = element.getAttribute('data-note-body')
    if (raw !== null) return Fragment.fromJSON(schema, JSON.parse(raw) as unknown)
  } catch {
    // Corpo ilegível: a nota chega vazia, sem derrubar a colagem.
  }
  return Fragment.from(schema.nodes['paragraph']!.create())
}

export const NoteRef = Node.create<NoteRefOptions>({
  name: 'noteRef',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  content: 'block+',

  addOptions() {
    return { notes: undefined }
  },

  addAttributes() {
    return {
      kind: {
        default: NoteKind.Footnote,
        parseHTML: (element) =>
          element.getAttribute('data-kind') === NoteKind.Endnote ? NoteKind.Endnote : NoteKind.Footnote,
        renderHTML: (attributes) => ({ 'data-kind': String(attributes['kind']) }),
      },
      nid: {
        default: null,
        parseHTML: (element) => element.getAttribute('data-nid'),
        renderHTML: (attributes) =>
          typeof attributes['nid'] === 'string' ? { 'data-nid': attributes['nid'] } : {},
      },
      mark: {
        default: null,
        parseHTML: (element) => element.getAttribute('data-mark'),
        renderHTML: (attributes) =>
          typeof attributes['mark'] === 'string' ? { 'data-mark': attributes['mark'] } : {},
      },
    }
  },

  parseHTML() {
    return [
      {
        tag: 'sup[data-note-ref]',
        // Acima do sobrescrito, que também reconhece o `<sup>` e levaria a nota embora.
        priority: 100,
        // Num atributo: `<p>` dentro de `<p>` faria o analisador partir o parágrafo.
        getContent: (element, schema) => contentFromJson(element as HTMLElement, schema),
      },
    ]
  },

  renderHTML({ node, HTMLAttributes }) {
    return [
      'sup',
      {
        ...HTMLAttributes,
        'data-note-ref': '',
        class: 'note-ref',
        'data-note-body': JSON.stringify(node.content.toJSON()),
      },
      markOf(node) ?? '',
    ]
  },

  addNodeView() {
    return noteRefView(this.editor)
  },

  renderText({ node }) {
    return markOf(node) ?? ''
  },

  addProseMirrorPlugins() {
    const notes = (): DocumentNotes | undefined => this.options.notes?.()
    return [
      new Plugin<NoteRefState>({
        key: noteRefKey,
        state: {
          init: (_config, state) => noteRefState(state.doc, notes(), []),
          apply: (transaction, previous, _old, state) => {
            const pages = transaction.getMeta(noteRefKey) as readonly (number | undefined)[] | undefined
            if (!transaction.docChanged && pages === undefined) return previous
            return noteRefState(state.doc, notes(), pages ?? previous.pages)
          },
        },
        props: {
          decorations: (state) => noteRefKey.getState(state)?.decorations ?? null,
          transformPasted: (slice, view) =>
            withoutRepeatedNotes(slice, view.state.doc, view.dragging?.move === true),
        },
      }),
    ]
  },
})

/** A referência tem o corpo dentro: sem isto o título com nota iria ao sumário com o texto dela. */
export function textBetweenWithoutNotes(
  node: ProseMirrorNode,
  from: number,
  to: number,
  blockSeparator?: string,
  leafText: (leaf: ProseMirrorNode) => string = () => '',
): string {
  let text = ''
  let first = true
  node.nodesBetween(from, to, (child, pos) => {
    if (child.type.name === 'noteRef') return false
    const own = child.isText
      ? (child.text ?? '').slice(Math.max(from, pos) - pos, to - pos)
      : child.isLeaf
        ? leafText(child)
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

/** `labels` são os da tela para esta folha, na mesma ordem (ver `print-source.ts`). */
export function numberNotesForPrint(holder: HTMLElement, labels: readonly string[]): void {
  let index = 0
  for (const element of holder.querySelectorAll<HTMLElement>('sup[data-note-ref]')) {
    element.removeAttribute('data-note-body')
    const label = labels[index]
    index += 1
    if (element.hasAttribute('data-mark') || label === undefined) continue
    element.textContent = label
  }
}
