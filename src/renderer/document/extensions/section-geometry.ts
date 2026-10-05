import { Extension } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { SectionSetup } from '@services/document/model.js'
import { blockSections, sectionBreakIn, type SectionBlock } from '@services/document/sections.js'

/**
 * A block from another section wraps lines at its sheet's width and starts at its margin. A
 * decoration, like the gaps (`pagination.ts`): `left` shifts without touching the indent, and a
 * negative right margin widens the box.
 */
export interface SectionBox {
  readonly shiftPx: number
  /** Negative means wider. */
  readonly narrowerPx: number
}

interface SectionGeometryState {
  readonly boxes: readonly SectionBox[]
  readonly declared: readonly SectionSetup[]
  readonly decorations: DecorationSet
}

export const sectionGeometryKey = new PluginKey<SectionGeometryState>('sectionGeometry')

/** Less than half a pixel is rounding. */
const TOLERANCE_PX = 0.5

function decorate(
  doc: ProseMirrorNode,
  boxes: readonly SectionBox[],
  declared: readonly SectionSetup[],
): DecorationSet {
  if (boxes.every((box) => Math.abs(box.shiftPx) < TOLERANCE_PX && Math.abs(box.narrowerPx) < TOLERANCE_PX)) {
    return DecorationSet.empty
  }

  const marks: (string | null)[] = []
  doc.forEach((block) => marks.push(sectionBreakIn(block as unknown as SectionBlock)))
  const sections = blockSections(marks, declared)

  const decorations: Decoration[] = []
  doc.forEach((block, offset, index) => {
    const box = boxes[sections[index] ?? boxes.length - 1]
    if (box === undefined) return
    if (Math.abs(box.shiftPx) < TOLERANCE_PX && Math.abs(box.narrowerPx) < TOLERANCE_PX) return
    decorations.push(
      Decoration.node(offset, offset + block.nodeSize, {
        style: `position:relative;left:${box.shiftPx}px;margin-right:${box.narrowerPx}px`,
        'data-section-box': '',
      }),
    )
  })
  return DecorationSet.create(doc, decorations)
}

export const SectionGeometry = Extension.create({
  name: 'sectionGeometry',

  addProseMirrorPlugins() {
    return [
      new Plugin<SectionGeometryState>({
        key: sectionGeometryKey,
        state: {
          init: () => ({ boxes: [], declared: [], decorations: DecorationSet.empty }),
          apply(transaction, current) {
            const next = transaction.getMeta(sectionGeometryKey) as
              Pick<SectionGeometryState, 'boxes' | 'declared'> | undefined
            if (next !== undefined) {
              return { ...next, decorations: decorate(transaction.doc, next.boxes, next.declared) }
            }
            // The mark comes and goes with editing: redone on every document change.
            if (!transaction.docChanged || current.boxes.length === 0) return current
            return { ...current, decorations: decorate(transaction.doc, current.boxes, current.declared) }
          },
        },
        props: {
          decorations: (state) => sectionGeometryKey.getState(state)?.decorations,
        },
      }),
    ]
  },
})

/** In `effectiveSections` order. Outside the history. */
export function setSectionBoxes(
  view: EditorView,
  boxes: readonly SectionBox[],
  declared: readonly SectionSetup[],
): void {
  const current = sectionGeometryKey.getState(view.state)
  if (
    current !== undefined &&
    current.declared === declared &&
    current.boxes.length === boxes.length &&
    current.boxes.every(
      (box, index) =>
        Math.abs(box.shiftPx - boxes[index]!.shiftPx) < TOLERANCE_PX &&
        Math.abs(box.narrowerPx - boxes[index]!.narrowerPx) < TOLERANCE_PX,
    )
  ) {
    return
  }
  view.dispatch(view.state.tr.setMeta(sectionGeometryKey, { boxes, declared }).setMeta('addToHistory', false))
}

/**
 * `bodySection` points to the entry standing in for the last section, as a document attribute so
 * undo carries it (`planSectionBreak`). The mark does not travel through paste: two marks with one
 * id would be two unordered sections, and Word does not copy the structure either.
 */
export const SectionMarks = Extension.create({
  name: 'sectionMarks',

  addGlobalAttributes() {
    return [
      {
        types: ['doc'],
        attributes: {
          bodySection: { default: null, rendered: false },
        },
      },
    ]
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        props: {
          transformPasted: (slice) => new Slice(withoutMarks(slice.content), slice.openStart, slice.openEnd),
        },
      }),
    ]
  },
})

function withoutMarks(fragment: Fragment): Fragment {
  const nodes: ProseMirrorNode[] = []
  fragment.forEach((node) => {
    const content = withoutMarks(node.content)
    const attrs =
      typeof node.attrs['sectionBreak'] === 'string' ? { ...node.attrs, sectionBreak: null } : node.attrs
    nodes.push(node.isText ? node : node.type.create(attrs, content, node.marks))
  })
  return Fragment.fromArray(nodes)
}
