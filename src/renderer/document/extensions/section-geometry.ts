import { Extension } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { SectionSetup } from '@services/document/model.js'
import { blockSections, sectionBreakIn, type SectionBlock } from '@services/document/sections.js'

/**
 * A caixa de texto de cada seção, na coluna única do editor.
 *
 * O editor é um fluxo só, com a largura da seção base (a última). Um bloco de
 * outra seção — a de paisagem, a de margens diferentes — precisa quebrar as
 * linhas na largura da sua folha e começar na margem dela, que fica noutro
 * lugar da pilha: a folha em paisagem é mais larga, e as folhas vão centradas.
 *
 * Decoração de nó, pela mesma razão dos vãos de página (`pagination.ts`): é
 * aparência sobre o nó, e não conteúdo. `left` desloca o bloco sem mexer na
 * margem esquerda dele, que já é o recuo do parágrafo e da lista; a margem
 * direita negativa alarga a caixa (ou a positiva a estreita) na medida exata
 * da diferença entre as duas colunas de texto.
 */
export interface SectionBox {
  /** Quanto a coluna da seção começa à direita da coluna base. */
  readonly shiftPx: number
  /** Quanto a coluna da seção é mais estreita que a base (negativo: mais larga). */
  readonly narrowerPx: number
}

interface SectionGeometryState {
  readonly boxes: readonly SectionBox[]
  readonly declared: readonly SectionSetup[]
  readonly decorations: DecorationSet
}

export const sectionGeometryKey = new PluginKey<SectionGeometryState>('sectionGeometry')

/** Menos de meio pixel de diferença é arredondamento, e não outra caixa. */
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
            // A marca de seção entra e sai com a edição, e o bloco muda de seção:
            // refeito a cada mudança do documento, que é uma volta pelos blocos
            // de primeiro nível.
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

/**
 * As caixas de cada seção, na ordem de `effectiveSections`.
 *
 * Fora do histórico: desfazer volta o que a pessoa escreveu, e não o lugar da
 * coluna.
 */
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
 * A estrutura das seções no texto.
 *
 * `bodySection`, no documento, aponta a entrada da biblioteca que faz as vezes
 * da última seção depois de uma quebra nova — é atributo do documento para que o
 * desfazer o leve junto com a marca (ver `planSectionBreak`).
 *
 * E a marca não viaja por colagem: o parágrafo colado (ou arrastado) levaria o
 * id de uma seção que já tem a sua marca, e duas marcas de um id são duas seções
 * que a gravação não sabe ordenar. A estrutura de seções não se copia, como no
 * Word ao colar texto dentro de uma seção.
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
