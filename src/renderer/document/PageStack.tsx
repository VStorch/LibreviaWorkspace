import { useMemo, type RefObject } from 'react'
import { EditorContent, type Editor } from '@tiptap/react'
import type { DocumentComment, PageSetup } from '@services/document/model.js'
import { pageDimensionsMm } from '@services/document/model.js'
import { sheetSetups } from '@services/document/sections.js'
import { mmToPx } from '@services/units.js'
import { CommentsPane } from './CommentsPane.js'
import { PaperSheet } from './PaperSheet.js'
import type { PlacedFloat } from './FloatingLayer.js'
import type { PageLayout } from './usePagination.js'
import { sheetBoxOf, type SheetGeometry } from './useSheetGeometry.js'
import type { SheetEditing } from './useSheetEditing.js'

export interface PageStackProps {
  readonly editor: Editor
  readonly layout: PageLayout
  readonly reading: boolean
  readonly zoom: number
  readonly geometry: SheetGeometry
  readonly effective: readonly PageSetup[]
  readonly floatsByPage: readonly PlacedFloat[][]
  readonly sheetEditing: SheetEditing | Record<string, never>
  readonly notePool: RefObject<HTMLDivElement | null>
  /** `null` com o painel fechado ou sem comentários. */
  readonly comments: {
    readonly list: readonly DocumentComment[]
    readonly outside: ReadonlySet<string>
  } | null
  /** O invólucro do zoom cresce para a rolagem chegar até a coluna dos comentários. */
  readonly zoomedWidthPx: number
}

/**
 * O `transform` do zoom não muda o espaço ocupado: o invólucro o reserva, para a
 * rolagem chegar ao fim. A paginação mede em 100 %.
 */
export function PageStack(props: PageStackProps): React.JSX.Element {
  const { editor, layout, reading, zoom, geometry, effective, notePool, comments, zoomedWidthPx } = props
  const page = effective.at(-1)!
  const { stackWidthPx, baseLeftPx, baseRightPx } = geometry
  return (
    <div
      className={`pages-zoom${reading ? ' pages-zoom--reading' : ''}`}
      style={
        reading
          ? undefined
          : { width: `${(zoomedWidthPx * zoom) / 100}px`, height: `${(layout.stackHeightPx * zoom) / 100}px` }
      }
    >
      <div
        className={`pages${reading ? ' pages--reading' : ''}`}
        data-zoom={reading ? 100 : zoom}
        style={
          reading
            ? undefined
            : {
                width: `${stackWidthPx}px`,
                height: `${layout.stackHeightPx}px`,
                ...(zoom === 100 ? {} : { transform: `scale(${zoom / 100})`, transformOrigin: 'top left' }),
              }
        }
      >
        {/* Fora do `contenteditable`: dentro, cada folha seria um nó
        selecionável. Sem folhas no modo de leitura, nem objetos ancorados,
        cuja posição é relativa a uma folha. */}
        {!reading && <SheetPapers {...props} />}

        <div
          className="pages__column"
          style={
            reading
              ? // A largura da leitura vem do CSS, e não das margens do documento.
                undefined
              : {
                  paddingTop: `${geometry.topInsetPx}px`,
                  paddingRight: `${baseRightPx}px`,
                  paddingLeft: `${baseLeftPx}px`,
                }
          }
        >
          <EditorContent editor={editor} />
        </div>

        {/* Os corpos de nota sem folha, na largura da coluna, onde a paginação
            os mede; no modo de leitura, aparecem aqui, depois do texto. */}
        <div
          ref={notePool}
          className={`note-pool${reading ? ' note-pool--reading' : ''}`}
          style={
            reading
              ? undefined
              : {
                  left: `${baseLeftPx}px`,
                  width: `${mmToPx(pageDimensionsMm(page).width - page.margins.left - page.margins.right)}px`,
                }
          }
        />

        {comments !== null && (
          <CommentsPane
            editor={editor}
            comments={comments.list}
            outside={comments.outside}
            leftPx={stackWidthPx}
          />
        )}
      </div>
    </div>
  )
}

/** O papel de cada folha e, nele, as faixas, que moram dentro da margem e não empurram o texto. */
function SheetPapers({
  editor,
  layout,
  geometry,
  effective,
  floatsByPage,
  sheetEditing,
}: PageStackProps): React.JSX.Element {
  const page = effective.at(-1)!
  const sheetSetupList = useMemo(() => sheetSetups(effective, layout.sheets), [effective, layout.sheets])
  const boxOf = (index: number): ReturnType<typeof sheetBoxOf> =>
    sheetBoxOf(effective, layout, geometry.stackWidthPx, index)
  return (
    <>
      {layout.sheetTops.map((top, index) => {
        const box = boxOf(index)
        return (
          <div
            key={top}
            className={`paper${(layout.sheetHeights[index] ?? 0) > box.heightPx + 1 ? ' paper--oversized' : ''}${layout.sheets[index]?.blank === true ? ' paper--blank' : ''}`}
            style={{
              top: `${top}px`,
              height: `${layout.sheetHeights[index] ?? box.heightPx}px`,
              left: `${box.leftPx}px`,
              width: `${box.widthPx}px`,
              right: 'auto',
            }}
            data-section={layout.sheets[index]?.section ?? 0}
            aria-hidden="true"
          >
            <span className="paper__number">{index + 1}</span>
          </div>
        )
      })}

      {layout.sheetTops.map((top, index) => {
        const setup = sheetSetupList[index] ?? { page, inSection: index + 1 }
        return (
          <PaperSheet
            key={`banda-${top}`}
            page={setup.page}
            pageNumber={setup.inSection}
            totalPages={layout.pages}
            topPx={top}
            leftPx={boxOf(index).leftPx}
            section={layout.sheets[index]?.section ?? 0}
            floats={floatsByPage[index] ?? []}
            columnLines={layout.columnLines.filter((line) => line.sheet === index)}
            noteAreas={layout.noteAreas.filter((area) => area.sheet === index)}
            schema={editor.schema}
            {...sheetEditing}
          />
        )
      })}
    </>
  )
}
