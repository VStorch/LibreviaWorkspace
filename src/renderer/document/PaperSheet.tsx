import type { Schema } from '@tiptap/pm/model'
import {
  bandForPage,
  bandInsetMm,
  hasBandContent,
  mmToPx,
  pageDimensionsMm,
  type DocumentNode,
  type PageSetup,
} from '@services/document/model.js'
import { bandFloatsOf } from '@services/document/floating.js'
import { FloatingLayer, type FloatSource, type PlacedFloat } from './FloatingLayer.js'
import { PageBand } from './PageBand.js'

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
  /** Os objetos ancorados em blocos que caíram nesta folha. */
  floats: readonly PlacedFloat[]
  schema: Schema
  /** Ausentes quando o documento está travado: aí nada recebe o cursor. */
  onEditFloat?: ((source: FloatSource, content: DocumentNode[]) => void) | undefined
  onEditBandPiece?: ((pid: string, text: string) => void) | undefined
  onEditBandBox?: ((bid: string, content: DocumentNode[]) => void) | undefined
}): React.JSX.Element {
  const { height } = pageDimensionsMm(page)
  const bandFloats = bandFloatsOf(page, pageNumber)
  const editFloat = onEditFloat === undefined ? {} : { onEdit: onEditFloat }

  return (
    <div className="paper-bands" style={{ top: `${topPx}px`, height: `${mmToPx(height)}px` }}>
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
            pageNumber={pageNumber}
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
    </div>
  )
}
