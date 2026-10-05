import type { Schema } from '@tiptap/pm/model'
import { pageDimensionsMm, type DocumentNode, type PageSetup } from '@services/document/model.js'
import { mmToPx } from '@services/units.js'
import { bandForPage, bandInsetMm, hasBandContent, pageLabel } from '@services/document/band.js'
import { bandFloatsOf } from '@services/document/floating.js'
import { FloatingLayer, type FloatSource, type PlacedFloat } from './FloatingLayer.js'
import { PageBand } from './PageBand.js'
import { NoteAreaView } from './NoteArea.js'
import type { NoteArea } from './usePagination.js'

export interface PaperSheetProps {
  readonly page: PageSetup
  /** From 1. */
  readonly pageNumber: number
  readonly totalPages: number
  readonly topPx: number
  /** The stack is as wide as the widest sheet. */
  readonly leftPx?: number
  /** Band heights are measured by it (`useBandHeights`). */
  readonly section?: number
  /** Sheet pixels. */
  readonly columnLines?: readonly {
    readonly leftPx: number
    readonly topPx: number
    readonly heightPx: number
  }[]
  readonly noteAreas?: readonly NoteArea[]
  readonly floats: readonly PlacedFloat[]
  readonly schema: Schema
  /** Absent when the document is locked: nothing takes the cursor. */
  readonly onEditFloat?: ((source: FloatSource, content: DocumentNode[]) => void) | undefined
  readonly onEditBandPiece?: ((pid: string, text: string) => void) | undefined
  readonly onEditBandBox?: ((bid: string, content: DocumentNode[]) => void) | undefined
}

/**
 * Outside the `contenteditable`: on paper these pieces live in the margin, without pushing the
 * text.
 */
export function PaperSheet({
  page,
  pageNumber,
  totalPages,
  topPx,
  leftPx = 0,
  section = 0,
  columnLines = [],
  noteAreas = [],
  floats,
  schema,
  onEditFloat,
  onEditBandPiece,
  onEditBandBox,
}: PaperSheetProps): React.JSX.Element {
  const { width, height } = pageDimensionsMm(page)
  const bandFloats = bandFloatsOf(page, pageNumber)
  const editFloat = onEditFloat === undefined ? {} : { onEdit: onEditFloat }

  return (
    <div
      className="paper-bands"
      data-section={section}
      style={{
        top: `${topPx}px`,
        height: `${mmToPx(height)}px`,
        left: `${leftPx}px`,
        width: `${mmToPx(width)}px`,
        right: 'auto',
      }}
    >
      <FloatingLayer objects={floats} page={page} schema={schema} behind {...editFloat} />
      {(['header', 'footer'] as const).map((kind) => {
        // Word's order: the title page beats parity, and parity beats the default.
        const band = bandForPage(page, pageNumber, kind)
        if (!hasBandContent(band)) return null

        return (
          <PageBand
            key={kind}
            band={band}
            kind={kind}
            pageLabel={pageLabel(page, pageNumber)}
            totalPages={totalPages}
            insetPx={mmToPx(bandInsetMm(page))}
            offsetPx={mmToPx(kind === 'header' ? page.headerDistanceMm : page.footerDistanceMm)}
            {...(onEditBandPiece === undefined ? {} : { onEdit: onEditBandPiece })}
          />
        )
      })}
      {/* In a layer of its own, above the text column, which covers the margin and would catch the click. */}
      {(['behind', 'front'] as const).map((where) => (
        <FloatingLayer
          key={where}
          objects={bandFloats}
          page={page}
          schema={schema}
          behind={where === 'behind'}
          variant="band"
          {...(onEditBandBox === undefined ? {} : { onEditBand: onEditBandBox })}
        />
      ))}
      <FloatingLayer objects={floats} page={page} schema={schema} behind={false} {...editFloat} />
      {columnLines.map((line) => (
        <div
          key={`${line.leftPx}:${line.topPx}`}
          className="paper-column-line"
          style={{ left: `${line.leftPx}px`, top: `${line.topPx}px`, height: `${line.heightPx}px` }}
        />
      ))}
      {noteAreas.map((area) => (
        <NoteAreaView key={area.kind} area={area} />
      ))}
    </div>
  )
}
