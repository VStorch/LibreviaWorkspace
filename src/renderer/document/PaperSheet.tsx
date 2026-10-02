import type { Schema } from '@tiptap/pm/model'
import { mmToPx, pageDimensionsMm, type DocumentNode, type PageSetup } from '@services/document/model.js'
import { bandForPage, bandInsetMm, hasBandContent, pageLabel } from '@services/document/band.js'
import { bandFloatsOf } from '@services/document/floating.js'
import { FloatingLayer, type FloatSource, type PlacedFloat } from './FloatingLayer.js'
import { PageBand } from './PageBand.js'
import { NoteAreaView } from './NoteArea.js'
import type { NoteArea } from './usePagination.js'

/**
 * O que se desenha por cima de uma folha: cabeçalho, rodapé e objetos ancorados.
 *
 * Fica **fora** do `contenteditable`, na camada das folhas: no papel essas
 * peças moram dentro da margem, e ali não empurram o texto nem entram na
 * seleção. A faixa é uma só no arquivo mesmo aparecendo em todas as folhas.
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
}: {
  page: PageSetup
  /** Qual folha esta é, começando em 1. */
  pageNumber: number
  totalPages: number
  /** Onde a folha começa na pilha desenhada. */
  topPx: number
  /** Onde a folha começa na horizontal: a pilha tem a largura da folha mais larga. */
  leftPx?: number
  /** A seção da folha — é por ela que a altura das faixas é medida (`useBandHeights`). */
  section?: number
  /** As linhas entre colunas desta folha, em pixels da folha. */
  columnLines?: readonly { readonly leftPx: number; readonly topPx: number; readonly heightPx: number }[]
  /** As notas de rodapé e de fim desta folha (M11). */
  noteAreas?: readonly NoteArea[]
  /** Os objetos ancorados em blocos que caíram nesta folha. */
  floats: readonly PlacedFloat[]
  schema: Schema
  /** Ausentes quando o documento está travado: aí nada recebe o cursor. */
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
        // A capa manda sobre a paridade, e a paridade sobre o padrão — a ordem
        // do Word. Documento sem primeira página distinta cai no padrão, e nada
        // muda para ele.
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

      {/* Os objetos da faixa vêm por último e numa camada própria: a faixa
          repete em toda folha, mora na margem e não disputa espaço com o corpo.
          É a mesma razão pela qual a faixa fica acima da coluna de texto — o
          retângulo da coluna cobre a margem inteira e apanhava o clique
          destinado à caixa. */}
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
