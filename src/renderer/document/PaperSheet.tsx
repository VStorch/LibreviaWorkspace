import type { Schema } from '@tiptap/pm/model'
import { pageDimensionsMm, type DocumentNode, type PageSetup } from '@services/document/model.js'
import { mmToPx } from '@services/units.js'
import { bandForPage, bandInsetMm, hasBandContent, pageLabel } from '@services/document/band.js'
import { bandFloatsOf } from '@services/document/floating.js'
import { FloatingLayer, type FloatSource, type PlacedFloat } from './FloatingLayer.js'
import { PageBand } from './PageBand.js'
import { NoteAreaView } from './NoteArea.js'
import type { NoteArea } from './usePagination.js'

/** Fora do `contenteditable`: no papel estas peças moram na margem, sem empurrar o texto. */
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
}: {
  page: PageSetup
  /** A partir de 1. */
  pageNumber: number
  totalPages: number
  topPx: number
  /** A pilha tem a largura da folha mais larga. */
  leftPx?: number
  /** É por ela que a altura das faixas é medida (`useBandHeights`). */
  section?: number
  /** Em pixels da folha. */
  columnLines?: readonly { readonly leftPx: number; readonly topPx: number; readonly heightPx: number }[]
  noteAreas?: readonly NoteArea[]
  floats: readonly PlacedFloat[]
  schema: Schema
  /** Ausentes no documento travado: nada recebe o cursor. */
  onEditFloat?: ((source: FloatSource, content: DocumentNode[]) => void) | undefined
  onEditBandPiece?: ((pid: string, text: string) => void) | undefined
  onEditBandBox?: ((bid: string, content: DocumentNode[]) => void) | undefined
}): React.JSX.Element {
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
        // A ordem do Word: a capa manda sobre a paridade, e a paridade sobre o padrão.
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

      {/* Numa camada própria, acima da coluna de texto, que cobre a margem e apanharia o clique. */}
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
