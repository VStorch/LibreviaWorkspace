import { useCallback, useMemo } from 'react'
import type { Editor } from '@tiptap/react'
import type { DocumentNode, PageSetup } from '@services/document/model.js'
import { editBandFloat, editBandPiece } from '@services/document/band.js'
import { floatsOf } from '@services/document/floating.js'
import { pxToMm } from '@services/units.js'
import { useWorkspace } from '../state/workspace.js'
import type { FloatSource, PlacedFloat } from './FloatingLayer.js'
import type { PageLayout } from './usePagination.js'

/** A posição depende da folha em que o parágrafo âncora caiu. */
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
        // É pela posição do bloco que o texto da caixa volta ao atributo de onde saiu.
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

/** Vazio no somente leitura: a folha não oferece edição. */
export function useSheetEditing(
  editor: Editor | null,
  readOnly: boolean,
): SheetEditing | Record<string, never> {
  const setPage = useWorkspace((state) => state.setPage)
  const setSections = useWorkspace((state) => state.setSections)

  /** Uma transação comum: entra no histórico, suja o documento e chega ao gravador pelo `getJSON()`. */
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

  /** A faixa mora em `page`, e não no documento do editor: por isso `setPage`, e não transação. */
  const editAllSections = useCallback(
    (change: <T extends PageSetup>(section: T) => T) => {
      // Da loja: várias peças podem sair do foco em sequência.
      const state = useWorkspace.getState()
      const page = change(state.page)
      if (page !== state.page) setPage(page)
      const sections = state.sections.map(change)
      if (sections.some((section, index) => section !== state.sections[index])) setSections(sections)
    },
    [setPage, setSections],
  )

  // A quebra de seção copia as referências, e a mesma parte mora nas duas seções.
  const onEditBandPiece = useCallback(
    (pid: string, text: string) => {
      if (!readOnly) editAllSections((section) => editBandPiece(section, pid, text))
    },
    [readOnly, editAllSections],
  )

  /** A caixa vem inteira: digitar dentro dela abre e fecha parágrafos. */
  const onEditBandBox = useCallback(
    (bid: string, content: DocumentNode[]) => {
      if (!readOnly) editAllSections((section) => editBandFloat(section, bid, content))
    },
    [readOnly, editAllSections],
  )

  return readOnly ? {} : { onEditFloat, onEditBandPiece, onEditBandBox }
}
