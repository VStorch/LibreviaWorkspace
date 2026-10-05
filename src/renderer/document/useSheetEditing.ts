import { useCallback, useMemo } from 'react'
import type { Editor } from '@tiptap/react'
import type { DocumentNode, PageSetup } from '@services/document/model.js'
import { editBandFloat, editBandPiece } from '@services/document/band.js'
import { floatsOf } from '@services/document/floating.js'
import { pxToMm } from '@services/units.js'
import { useWorkspace } from '../state/workspace.js'
import type { FloatSource, PlacedFloat } from './FloatingLayer.js'
import type { PageLayout } from './usePagination.js'

/** The position depends on the sheet the anchor paragraph fell on. */
export function useFloatsByPage(
  editor: Editor | null,
  layout: PageLayout,
  revision: number,
): PlacedFloat[][] {
  return useMemo(() => {
    const pages: PlacedFloat[][] = Array.from({ length: layout.pages }, () => [])
    if (editor === null) return pages

    let index = 0
    editor.state.doc.forEach((node, pos) => {
      const anchor = layout.anchors[index]
      index += 1
      if (anchor === undefined) return

      const sheet = pages[anchor.pageIndex]
      if (sheet === undefined) return

      let slot = 0
      for (const object of floatsOf(node.attrs)) {
        // The block position is how the box text goes back to the attribute it came from.
        sheet.push({ object, anchorTopMm: pxToMm(anchor.topPx), source: { pos, index: slot } })
        slot += 1
      }
    })

    return pages
  }, [editor, layout, revision])
}

export interface SheetEditing {
  readonly onEditFloat: (source: FloatSource, content: DocumentNode[]) => void
  readonly onEditBandPiece: (pid: string, text: string) => void
  readonly onEditBandBox: (bid: string, content: DocumentNode[]) => void
}

/** Empty when read-only: the sheet offers no editing. */
export function useSheetEditing(
  editor: Editor | null,
  readOnly: boolean,
): SheetEditing | Record<string, never> {
  const setPage = useWorkspace((state) => state.setPage)
  const setSections = useWorkspace((state) => state.setSections)

  /**
   * An ordinary transaction: it enters history, marks the document dirty and reaches the writer
   * through `getJSON()`.
   */
  const onEditFloat = useCallback(
    (source: FloatSource, content: DocumentNode[]) => {
      if (editor === null || readOnly) return

      const node = editor.state.doc.nodeAt(source.pos)
      if (node === null) return

      const floats = node.attrs['floats']
      if (!Array.isArray(floats)) return

      const object = floats[source.index] as { content?: unknown } | undefined
      if (object === undefined) return
      if (JSON.stringify(object.content ?? []) === JSON.stringify(content)) return

      const updated = floats.map((item, index) => (index === source.index ? { ...object, content } : item))
      editor.view.dispatch(editor.state.tr.setNodeAttribute(source.pos, 'floats', updated))
    },
    [editor, readOnly],
  )

  /** The band lives in `page`, not in the editor document: hence `setPage`, not a transaction. */
  const editAllSections = useCallback(
    (change: <T extends PageSetup>(section: T) => T) => {
      // From the store: several pieces may lose focus in a row.
      const state = useWorkspace.getState()
      const page = change(state.page)
      if (page !== state.page) setPage(page)
      const sections = state.sections.map(change)
      if (sections.some((section, index) => section !== state.sections[index])) setSections(sections)
    },
    [setPage, setSections],
  )

  // A section break copies the references, and the same part lives in both sections.
  const onEditBandPiece = useCallback(
    (pid: string, text: string) => {
      if (!readOnly) editAllSections((section) => editBandPiece(section, pid, text))
    },
    [readOnly, editAllSections],
  )

  /** The whole box comes in: typing inside it opens and closes paragraphs. */
  const onEditBandBox = useCallback(
    (bid: string, content: DocumentNode[]) => {
      if (!readOnly) editAllSections((section) => editBandFloat(section, bid, content))
    },
    [readOnly, editAllSections],
  )

  return readOnly ? {} : { onEditFloat, onEditBandPiece, onEditBandBox }
}
