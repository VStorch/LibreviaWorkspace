import { useMemo, type MouseEvent, type RefObject } from 'react'
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
  /** `null` with the pane closed or without comments. */
  readonly comments: {
    readonly list: readonly DocumentComment[]
    readonly outside: ReadonlySet<string>
  } | null
  /** The zoom wrapper grows so scrolling reaches the comments column. */
  readonly zoomedWidthPx: number
}

/**
 * The zoom `transform` does not change the space taken: the wrapper reserves it so scrolling
 * reaches the end. Pagination measures at 100 %.
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
        onMouseDown={(event) => placeCaretNearClick(editor, event)}
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
        {/* Outside the `contenteditable`: inside, each sheet would be a selectable node. No sheets in reading
        mode, and no anchored objects, whose position is relative to a sheet. */}
        {!reading && <SheetPapers {...props} />}
        <div
          className="pages__column"
          style={
            reading
              ? // Reading width comes from CSS, not from the document margins.
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
        {/* Note bodies without a sheet, at column width, where pagination measures them; in reading mode they
        show here, after the text. */}
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

/**
 * The paper is drawn behind the text and the editable area ends at the last line, so a click lower on
 * the sheet would reach no text, as in a new document.
 */
function placeCaretNearClick(editor: Editor, event: MouseEvent<HTMLDivElement>): void {
  const target = event.target as Element
  const onBareSheet = target === event.currentTarget || target.classList.contains('pages__column')
  if (event.button !== 0 || !onBareSheet || !editor.isEditable) return

  const text = editor.view.dom.getBoundingClientRect()
  const found = editor.view.posAtCoords({
    left: Math.min(Math.max(event.clientX, text.left), text.right - 1),
    top: Math.min(Math.max(event.clientY, text.top), text.bottom - 1),
  })
  if (found === null) return

  event.preventDefault()
  // Tiptap's focus() waits a frame; keys typed right after the click would land on the body.
  editor.view.focus()
  editor.commands.setTextSelection(found.pos)
}

/**
 * Each sheet's paper and, on it, the bands, which live inside the margin and do not push the text.
 */
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
