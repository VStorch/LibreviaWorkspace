import { Node } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode, type Schema } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { DocumentNotes } from '@services/document/model.js'
import { NoteKind, noteLabels } from '@services/document/notes.js'
import { sectionBreakIn, type SectionBlock } from '@services/document/sections.js'
import { noteRefView } from './note-view.js'

/**
 * An inline atomic node **with content**: the body has its own editor (`note-view.ts`), but travels
 * inside the reference, so copying copies the note and deleting deletes it. `nid` is the `w:id`;
 * one pasted beside the original loses it and gets its own note. The number is the reference's
 * order, drawn by a decoration.
 */

export interface NoteRefOptions {
  /** Consultada a cada conta. */
  readonly notes: (() => DocumentNotes | undefined) | undefined
}

export const noteRefKey = new PluginKey<NoteRefState>('noteRef')

/** Without descending into any body. */
export function noteRefsOf(doc: ProseMirrorNode): Array<{ node: ProseMirrorNode; pos: number }> {
  const found: Array<{ node: ProseMirrorNode; pos: number }> = []
  doc.descendants((node, pos) => {
    if (node.type.name !== 'noteRef') return true
    found.push({ node, pos })
    return false
  })
  return found
}

/** `null` outside a note: that is how a command knows the selection is the body's. */
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

/**
 * Not one with its own mark: Word writes the mark in the body, and drawing it again would give
 * "**".
 */
export function drawsNoteNumber(node: ProseMirrorNode): boolean {
  return markOf(node) === null
}

/** By counting the section marks before it, which close the section like `w:sectPr`. */
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

/** `pages` is each one's sheet, for per-page restart. */
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

/** Paper and `NOTEREF` use the same ones. */
export function noteLabelsOf(state: EditorState): readonly string[] {
  // Without the plugin (a separately built state), the document's default count.
  return noteRefKey.getState(state)?.labels ?? noteRefLabels(state.doc)
}

export function notePagesOf(state: EditorState): readonly (number | undefined)[] {
  return noteRefKey.getState(state)?.pages ?? []
}

/** Comes from pagination once it settles; the transaction does not change the document. */
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
 * The file does not accept two references to the same note: a pasted one goes without `nid` and
 * gets its own note. Dragging is not pasting, and keeps the `nid`.
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
    // Unreadable body: the note arrives empty, without breaking the paste.
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
        // Above superscript, which also recognizes `<sup>` and would take the note away.
        priority: 100,
        // In an attribute: `<p>` inside `<p>` would make the parser split the paragraph.
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

/**
 * The reference has the body inside: without this a heading with a note would go to the table of
 * contents with the note text.
 */
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

/** `labels` are the screen's for this sheet, in the same order (see `print-source.ts`). */
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
