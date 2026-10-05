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
  /** Each sheet centered in the stack, as Word shows portrait and landscape together. */
  readonly stackWidthPx: number
  readonly baseLeftPx: number
  readonly baseRightPx: number
  /** The first sheet's top margin: a floor, as in Word. */
  readonly topInsetPx: number
}

export interface SheetGeometryInput {
  readonly editor: Editor | null
  /** With bands resolved; the last is the body's. */
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

  // Reading mode has no sheets, and nothing moves.
  useEffect(() => {
    if (editor === null) return
    const boxes = reading
      ? []
      : effective.map((section) => {
          const widthPx = mmToPx(pageDimensionsMm(section).width)
          const left = (stackWidthPx - widthPx) / 2 + mmToPx(section.margins.left)
          // Pagination decides which column, through a translation.
          const content = mmToPx(columnGeometry(section).widthMm)
          const base = stackWidthPx - baseLeftPx - baseRightPx
          return { shiftPx: left - baseLeftPx, narrowerPx: base - content }
        })
    setSectionBoxes(editor.view, boxes, sections)
  }, [editor, effective, sections, reading, stackWidthPx, baseLeftPx, baseRightPx])

  return { stackWidthPx, baseLeftPx, baseRightPx, topInsetPx: mmToPx(insets.top) }
}

/** Centered, with the paper of the section opening it. */
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
