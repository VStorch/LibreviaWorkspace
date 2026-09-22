import type { Editor } from '@tiptap/react'
import { DOMSerializer, Fragment, Node as ProseMirrorNode } from '@tiptap/pm/model'
import { pxToMm, type PageSetup } from '@services/document/model.js'
import { bandFloatsOf, floatsOf, type FloatingObject } from '@services/document/floating.js'
import type { PrintFloat, PrintPage } from '@services/document/print-pages.js'
import type { PageLayout, PageStart } from './usePagination.js'

/**
 * O documento recortado nas folhas que a tela mostra.
 *
 * Serializa **direto dos nós**, e não a partir de `getHTML()`. A diferença não
 * é de gosto: o leitor emite a quebra de página dentro do parágrafo quando o
 * Word a gravou assim (`w:br w:type="page"` no meio de um `w:r`), e um `<div>`
 * dentro de `<p>` faz o analisador de HTML fechar o parágrafo e desalojar o
 * `div`. Um documento de 15 nós virava 17 elementos, os índices deixavam de
 * casar, e o papel cortava em lugar diferente do da tela — que é exatamente o
 * defeito que este trabalho existe para acabar.
 *
 * Serializando o nó, o recorte cai sempre onde o paginador o pôs.
 */
export function splitIntoPages(editor: Editor, layout: PageLayout, page: PageSetup): PrintPage[] {
  const serializer = DOMSerializer.fromSchema(editor.schema)

  const blocks: ProseMirrorNode[] = []
  editor.state.doc.forEach((node: ProseMirrorNode) => blocks.push(node))

  const cuts: PageStart[] = [{ blockIndex: 0 }, ...layout.pageStarts, { blockIndex: blocks.length }]
  const pages: PrintPage[] = []

  for (let cut = 0; cut < cuts.length - 1; cut++) {
    const start = cuts[cut]!
    const end = cuts[cut + 1]!

    const holder = document.createElement('div')
    const fragments = slicePageBlocks(blocks, start, end)
    holder.appendChild(serializer.serializeFragment(Fragment.fromArray(fragments)))

    pages.push({
      number: pages.length + 1,
      html: holder.innerHTML,
      floats: [
        ...anchoredFloats(
          blocks,
          layout,
          start.blockIndex + (start.childIndex === undefined ? 0 : 1),
          end.blockIndex + (end.childIndex === undefined ? 0 : 1),
          editor,
        ),
        ...bandFloats(page, pages.length + 1, editor),
      ],
    })
  }

  return pages
}

/** Recorta linhas e itens sem duplicar conteúdo nem reiniciar listas numeradas. */
export function slicePageBlocks(
  blocks: readonly ProseMirrorNode[],
  start: PageStart,
  end: PageStart,
): ProseMirrorNode[] {
  const fragments: ProseMirrorNode[] = []
  for (let index = start.blockIndex; index <= end.blockIndex && index < blocks.length; index++) {
    const block = blocks[index]!
    const from = index === start.blockIndex ? (start.childIndex ?? 0) : 0
    const to = index === end.blockIndex ? (end.childIndex ?? 0) : block.childCount
    if (index === end.blockIndex && to === 0) break
    if (from === 0 && to === block.childCount) {
      fragments.push(block)
    } else {
      const children: ProseMirrorNode[] = []
      block.forEach((child, _offset, childIndex) => {
        if (childIndex >= from && childIndex < to) children.push(child)
      })
      const attrs =
        block.type.name === 'orderedList'
          ? { ...block.attrs, start: Number(block.attrs.start ?? 1) + from }
          : block.attrs
      fragments.push(block.type.create(attrs, Fragment.fromArray(children), block.marks))
    }
  }
  return fragments
}

/**
 * Os objetos ancorados nos blocos desta folha, com a caixa de texto já
 * serializada: o desenho do papel não conhece o schema do ProseMirror, e quem o
 * conhece é aqui.
 */
function anchoredFloats(
  blocks: readonly ProseMirrorNode[],
  layout: PageLayout,
  start: number,
  end: number,
  editor: Editor,
): PrintFloat[] {
  const floats: PrintFloat[] = []

  for (let index = start; index < end; index++) {
    const anchor = layout.anchors[index]
    const block = blocks[index]
    if (anchor === undefined || block === undefined) continue

    for (const object of floatsOf(block.attrs)) {
      floats.push({ object, anchorTopMm: pxToMm(anchor.topPx), ...contentHtmlOf(object, editor) })
    }
  }

  return floats
}

/** As caixas do cabeçalho e do rodapé, pela mesma razão: o HTML sai daqui. */
function bandFloats(page: PageSetup, pageNumber: number, editor: Editor): PrintFloat[] {
  return bandFloatsOf(page, pageNumber).map((item) => ({
    ...item,
    ...contentHtmlOf(item.object, editor),
  }))
}

/** O conteúdo de uma caixa de texto, em HTML, pelo serializador do editor. */
function contentHtmlOf(object: FloatingObject, editor: Editor): { contentHtml?: string } {
  if (object.kind !== 'text') return {}

  try {
    const nodes = (object.content ?? []).map((node) => ProseMirrorNode.fromJSON(editor.schema, node))
    const holder = document.createElement('div')
    holder.appendChild(DOMSerializer.fromSchema(editor.schema).serializeFragment(Fragment.fromArray(nodes)))
    return { contentHtml: holder.innerHTML }
  } catch {
    // Caixa que o schema não reconhece sai vazia em vez de derrubar a
    // exportação inteira.
    return { contentHtml: '' }
  }
}
