import { Extension } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { SectionSetup } from '@services/document/model.js'
import { blockSections, sectionBreakIn, type SectionBlock } from '@services/document/sections.js'

/**
 * O bloco de outra seção quebra as linhas na largura da folha dela e começa na
 * margem dela. Decoração, como os vãos (`pagination.ts`): `left` desloca sem
 * mexer no recuo, e a margem direita negativa alarga a caixa.
 */
export interface SectionBox {
  readonly shiftPx: number
  /** Negativo: mais larga. */
  readonly narrowerPx: number
}

interface SectionGeometryState {
  readonly boxes: readonly SectionBox[]
  readonly declared: readonly SectionSetup[]
  readonly decorations: DecorationSet
}

export const sectionGeometryKey = new PluginKey<SectionGeometryState>('sectionGeometry')

/** Menos de meio pixel é arredondamento. */
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
            // A marca entra e sai com a edição: refeito a cada mudança do documento.
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

/** Na ordem de `effectiveSections`. Fora do histórico. */
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
 * `bodySection` aponta a entrada que faz as vezes da última seção, como atributo
 * do documento para o desfazer levá-lo (`planSectionBreak`). A marca não viaja
 * por colagem: duas marcas de um id seriam duas seções sem ordem, e o Word também
 * não copia a estrutura.
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
