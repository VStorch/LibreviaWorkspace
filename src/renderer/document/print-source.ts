import type { Editor } from '@tiptap/react'
import { DOMSerializer, Fragment, Node as ProseMirrorNode } from '@tiptap/pm/model'
import { pxToMm, type PageSetup } from '@services/document/model.js'
import { bandFloatsOf, floatsOf, type FloatingObject } from '@services/document/floating.js'
import type { PrintFloat, PrintPage } from '@services/document/print-pages.js'
import { isInternalStart, type PageLayout, type PageStart } from './usePagination.js'

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
  const offsets: number[] = []
  editor.state.doc.forEach((node: ProseMirrorNode, offset: number) => {
    blocks.push(node)
    offsets.push(offset)
  })

  const cuts: PageStart[] = [{ blockIndex: 0 }, ...layout.pageStarts, { blockIndex: blocks.length }]
  const pages: PrintPage[] = []

  for (let cut = 0; cut < cuts.length - 1; cut++) {
    const start = cuts[cut]!
    const end = cuts[cut + 1]!

    const holder = document.createElement('div')
    const fragments = slicePageBlocks(blocks, start, end)
    holder.appendChild(serializer.serializeFragment(Fragment.fromArray(fragments)))
    // A mesma marca que a decoração põe na tela: o parágrafo da captura com
    // texto não ganha a linha vazia de 1lh.
    for (const paragraph of holder.querySelectorAll('p')) {
      if (
        paragraph.querySelector(':scope > img[data-anchored]') !== null &&
        (paragraph.textContent ?? '').trim() !== ''
      ) {
        paragraph.setAttribute('data-anchor-text', '')
      }
    }
    markSplitParagraphs(
      holder,
      start,
      end,
      end.offset !== undefined && isJustified(editor, offsets[end.blockIndex]),
    )

    pages.push({
      number: pages.length + 1,
      html: holder.innerHTML,
      floats: [
        ...anchoredFloats(
          blocks,
          layout,
          start.blockIndex + (isInternalStart(start) ? 1 : 0),
          end.blockIndex + (isInternalStart(end) ? 1 : 0),
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
    // Parágrafo cortado entre linhas: o recorte é do conteúdo, no caractere
    // em que a tela pôs o espaçador.
    const textFrom = index === start.blockIndex ? start.offset : undefined
    const textTo = index === end.blockIndex ? end.offset : undefined
    if (block.isTextblock && (textFrom !== undefined || textTo !== undefined)) {
      if (textTo === 0) break
      fragments.push(block.cut(textFrom ?? 0, textTo ?? block.content.size))
      continue
    }
    const from = index === start.blockIndex ? (start.childIndex ?? 0) : 0
    const to = index === end.blockIndex ? (end.childIndex ?? 0) : block.childCount
    if (index === end.blockIndex && to === 0) break
    if (from === 0 && to === block.childCount) {
      fragments.push(block)
    } else {
      const children: ProseMirrorNode[] = []
      block.forEach((child, _offset, childIndex) => {
        // As linhas de cabeçalho voltam no alto da folha em que a tabela
        // continua, como a tela as desenha.
        const repeated =
          index === start.blockIndex && start.repeatHeader === true && isHeaderRow(block, childIndex)
        if (repeated || (childIndex >= from && childIndex < to)) children.push(child)
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

/** Linha de cabeçalho: está no começo da tabela e só tem células `tableHeader`. */
function isHeaderRow(table: ProseMirrorNode, rowIndex: number): boolean {
  for (let index = 0; index <= rowIndex; index++) {
    const row = table.maybeChild(index)
    if (row === null || row.childCount === 0) return false
    let header = true
    row.forEach((cell) => {
      if (cell.type.name !== 'tableHeader') header = false
    })
    if (!header) return false
  }
  return true
}

/**
 * A costura do parágrafo que a folha cortou entre linhas.
 *
 * A parte de cima perde o espaço depois, e a de baixo o espaço antes e o recuo
 * da primeira linha: na tela os dois pedaços são um parágrafo só, e só a
 * primeira linha dele tem recuo. A última linha da parte de cima era uma linha
 * do meio, e no parágrafo justificado continua justificada — sem isto ela
 * sairia alinhada à esquerda, como última linha que o papel acha que é.
 *
 * O objeto ancorado vai com o pedaço de cima, que é onde o parágrafo começa;
 * `anchoredFloats` já conta o bloco na folha em que ele abre.
 */
function markSplitParagraphs(
  holder: HTMLElement,
  start: PageStart,
  end: PageStart,
  justified: boolean,
): void {
  const first = holder.firstElementChild
  if (start.offset !== undefined && first instanceof HTMLElement) {
    first.style.marginTop = '0'
    first.style.paddingTop = '0'
    first.style.textIndent = '0'
    first.dataset.continued = 'from'
  }
  const last = holder.lastElementChild
  if (end.offset !== undefined && last instanceof HTMLElement) {
    last.style.marginBottom = '0'
    last.style.paddingBottom = '0'
    if (justified) last.style.textAlignLast = 'justify'
    last.dataset.continued = 'to'
  }
}

/** O alinhamento que se vê, que pode vir do estilo e não do nó. */
function isJustified(editor: Editor, offset: number | undefined): boolean {
  if (offset === undefined) return false
  const dom = editor.view.nodeDOM(offset)
  return dom instanceof HTMLElement && getComputedStyle(dom).textAlign === 'justify'
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
