import { useEffect } from 'react'
import type { Editor } from '@tiptap/react'
import {
  contentInsetsMm,
  pageDimensionsMm,
  type PageSetup,
  type SectionSetup,
} from '@services/document/model.js'
import { NO_BANDS, type BandHeights } from '@services/document/band.js'
import { columnGeometry } from '@services/document/sections.js'
import { mmToPx } from '@services/units.js'
import { setSectionBoxes } from './extensions/section-geometry.js'
import type { PageLayout } from './usePagination.js'

export interface SheetGeometry {
  /** Cada folha centrada na pilha, como o Word mostra retrato e paisagem juntos. */
  readonly stackWidthPx: number
  readonly baseLeftPx: number
  readonly baseRightPx: number
  /** A margem de cima da primeira folha: é um piso, como no Word. */
  readonly topInsetPx: number
}

export interface SheetGeometryInput {
  readonly editor: Editor | null
  /** Com as faixas resolvidas; a última é a do corpo. */
  readonly effective: readonly PageSetup[]
  readonly sections: readonly SectionSetup[]
  readonly layout: PageLayout
  readonly bands: readonly BandHeights[]
  readonly reading: boolean
}

export function useSheetGeometry({
  editor,
  effective,
  sections,
  layout,
  bands,
  reading,
}: SheetGeometryInput): SheetGeometry {
  const page = effective.at(-1)!
  const pageWidthPx = mmToPx(pageDimensionsMm(page).width)
  const stackWidthPx = Math.max(layout.stackWidthPx, pageWidthPx)
  const firstSection = effective[layout.sheets[0]?.section ?? 0] ?? page
  const insets = contentInsetsMm(firstSection, bands[layout.sheets[0]?.section ?? 0] ?? NO_BANDS)
  const baseLeftPx = (stackWidthPx - pageWidthPx) / 2 + mmToPx(page.margins.left)
  const baseRightPx = (stackWidthPx - pageWidthPx) / 2 + mmToPx(page.margins.right)

  // No modo de leitura não há folha, e nada se desloca.
  useEffect(() => {
    if (editor === null) return
    const boxes = reading
      ? []
      : effective.map((section) => {
          const widthPx = mmToPx(pageDimensionsMm(section).width)
          const left = (stackWidthPx - widthPx) / 2 + mmToPx(section.margins.left)
          // Qual coluna, quem decide é a paginação, por translação.
          const content = mmToPx(columnGeometry(section).widthMm)
          const base = stackWidthPx - baseLeftPx - baseRightPx
          return { shiftPx: left - baseLeftPx, narrowerPx: base - content }
        })
    setSectionBoxes(editor.view, boxes, sections)
  }, [editor, effective, sections, reading, stackWidthPx, baseLeftPx, baseRightPx])

  return { stackWidthPx, baseLeftPx, baseRightPx, topInsetPx: mmToPx(insets.top) }
}

/** A folha `index` na pilha: centrada, com o papel da seção que a abre. */
export function sheetBoxOf(
  effective: readonly PageSetup[],
  layout: PageLayout,
  stackWidthPx: number,
  index: number,
): { leftPx: number; widthPx: number; heightPx: number } {
  const setup = effective[layout.sheets[index]?.section ?? effective.length - 1] ?? effective.at(-1)!
  const { width, height } = pageDimensionsMm(setup)
  return { leftPx: (stackWidthPx - mmToPx(width)) / 2, widthPx: mmToPx(width), heightPx: mmToPx(height) }
}
