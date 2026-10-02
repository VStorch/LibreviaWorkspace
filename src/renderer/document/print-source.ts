import type { Editor } from '@tiptap/react'
import { DOMSerializer, Fragment, Node as ProseMirrorNode } from '@tiptap/pm/model'
import { pxToMm, type PageSetup } from '@services/document/model.js'
import { bandFloatsOf, floatsOf, type FloatingObject } from '@services/document/floating.js'
import type { PrintFloat, PrintNoteArea, PrintPage } from '@services/document/print-pages.js'
import { drawListsForPrint } from './extensions/list-numbering.js'
import { sheetSetups } from '@services/document/sections.js'
import { RevisionView } from '@shared/types.js'
import { DELETION, INSERTION } from './extensions/track-changes.js'
import { isHiddenBlock, isHiddenInline, revisionViewOf } from './extensions/revision-view.js'
import { drawsNoteNumber, noteLabelsOf, noteRefsOf, numberNotesForPrint } from './extensions/note-ref.js'
import { NOTE_NUMBER_CLASS } from './extensions/note-view.js'
import {
  NOTE_SEPARATOR_PX,
  collapsed,
  drawnSheet,
  isInternalStart,
  type PageLayout,
  type PageStart,
} from './usePagination.js'

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
export function splitIntoPages(
  editor: Editor,
  layout: PageLayout,
  sections: readonly PageSetup[],
): PrintPage[] {
  const serializer = DOMSerializer.fromSchema(editor.schema)

  // Com a numeração das listas gravada nos nós: o serializador não vê as
  // decorações que a desenham na tela.
  const blocks = drawListsForPrint(editor.state.doc)
  // As alterações saem como a janela as mostra (Revisão → Mostrar).
  const view = revisionViewOf(editor.state)
  const offsets: number[] = []
  editor.state.doc.forEach((_node: ProseMirrorNode, offset: number) => {
    offsets.push(offset)
  })

  const cuts: PageStart[] = [{ blockIndex: 0 }, ...layout.pageStarts, { blockIndex: blocks.length }]
  const pages: PrintPage[] = []
  // O número das notas (M11) é decoração na tela; no papel ele é escrito aqui,
  // com os rótulos da tela — os reinícios por folha e por seção já contados.
  const screenLabels = noteLabelsOf(editor.state)
  const labelOf = new Map<ProseMirrorNode, string>()
  noteRefsOf(editor.state.doc).forEach(({ node }, index) => {
    const label = screenLabels[index]
    if (label !== undefined) labelOf.set(node, label)
  })
  // A configuração de cada folha desenhada: papel, faixas e número da seção dela.
  const setups = sheetSetups(sections, layout.sheets)
  const setupOf = (drawn: number): { setup: PageSetup; inSection: number } => {
    const found = setups[drawn]
    return found === undefined
      ? { setup: sections.at(-1)!, inSection: drawn + 1 }
      : { setup: found.page, inSection: found.inSection }
  }

  for (let cut = 0; cut < cuts.length - 1; cut++) {
    const start = cuts[cut]!
    const end = cuts[cut + 1]!

    // As folhas em branco que a seção par ou ímpar pediu antes desta: só a
    // faixa, como no Word.
    const drawn = drawnSheet(layout, cut)
    while (pages.length < drawn) {
      const blank = setupOf(pages.length)
      pages.push({
        number: pages.length + 1,
        html: '',
        floats: bandFloats(blank.setup, blank.inSection, editor),
        ...blank,
        blank: true,
      })
    }
    const sheet = setupOf(pages.length)

    const holder = document.createElement('div')
    // O recorte primeiro, nos índices da tela; o modo depois, que não os mexe.
    const fragments = blocksForView(slicePageBlocks(blocks, start, end), view)
    holder.appendChild(serializer.serializeFragment(Fragment.fromArray(fragments)))
    // Os comentários não vão ao papel, como no Word com a marcação desligada: o
    // painel e o realce são da tela (o realce é decoração, que o serializador não
    // vê), e as pontas saem aqui — vazias, mas são marcação de comentário.
    for (const anchor of holder.querySelectorAll('[data-comment-start], [data-comment-end]')) anchor.remove()
    numberNotesForPrint(holder, printedNoteLabels(fragments, labelOf))
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
    placeColumns(holder, start, end, layout)
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
        ...bandFloats(sheet.setup, sheet.inSection, editor),
      ],
      notes: notesForPrint(editor, layout, pages.length, view, screenLabels),
      columnLines: layout.columnLines
        .filter((line) => line.sheet === pages.length)
        .map((line) => ({
          leftMm: pxToMm(line.leftPx),
          topMm: pxToMm(line.topPx),
          heightMm: pxToMm(line.heightPx),
        })),
      ...sheet,
    })
  }

  return pages
}

/**
 * As notas da folha desenhada `sheet` (M11): as áreas que a tela pôs nela, com o
 * corpo de cada nota serializado do nó e o número escrito no começo — no papel
 * não há decoração para desenhá-lo.
 */
function notesForPrint(
  editor: Editor,
  layout: PageLayout,
  sheet: number,
  view: RevisionView,
  labels: readonly string[],
): PrintNoteArea[] {
  const areas = layout.noteAreas.filter((area) => area.sheet === sheet)
  if (areas.length === 0) return []
  const serializer = DOMSerializer.fromSchema(editor.schema)
  const refs = noteRefsOf(editor.state.doc)
  return areas.map((area) => ({
    topMm: pxToMm(area.topPx),
    leftMm: pxToMm(area.leftPx),
    widthMm: pxToMm(area.widthPx),
    separator: area.separator,
    separatorMm: pxToMm(NOTE_SEPARATOR_PX),
    items: area.items.flatMap((item) => {
      const reference = refs[item.index]
      if (reference === undefined) return []
      const children: ProseMirrorNode[] = []
      reference.node.forEach((child) => children.push(child))
      const holder = document.createElement('div')
      holder.appendChild(serializer.serializeFragment(Fragment.fromArray(blocksForView(children, view))))
      const first = holder.firstElementChild
      if (drawsNoteNumber(reference.node) && first !== null && /^(P|H[1-6])$/.test(first.tagName)) {
        const number = document.createElement('span')
        number.className = NOTE_NUMBER_CLASS
        number.textContent = labels[item.index] ?? ''
        first.prepend(number)
      }
      return [{ html: holder.innerHTML, clipTopMm: pxToMm(item.clipTopPx), heightMm: pxToMm(item.heightPx) }]
    }),
  }))
}

/**
 * Os blocos em coluna saem no papel como na tela (M9): a largura de uma coluna,
 * o lado da coluna dela e o desvio vertical do primeiro de cada coluna — os
 * mesmos números que a paginação da tela produziu.
 */
function placeColumns(holder: HTMLElement, start: PageStart, end: PageStart, layout: PageLayout): void {
  if (layout.columnMoves.length === 0) return
  const moves = new Map(layout.columnMoves.map((move) => [move.blockIndex, move]))
  const last = isInternalStart(end) ? end.blockIndex : end.blockIndex - 1
  const children = Array.from(holder.children)
  for (let index = start.blockIndex; index <= last; index++) {
    const element = children[index - start.blockIndex]
    const move = moves.get(index)
    if (!(element instanceof HTMLElement) || move === undefined) continue
    if (move.narrowerPx !== 0) element.style.marginRight = `${move.narrowerPx}px`
    if (move.dx !== 0) element.style.transform = `translateX(${move.dx}px)`
    // O primeiro da folha não sobe: a folha nova já o põe no topo.
    if (move.lift !== 0 && index !== start.blockIndex) {
      element.style.marginTop = `${collapsed(move.natural + move.lift, move.collapse)}px`
    }
  }
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

/**
 * Os blocos como o modo de mostrar as alterações os vê. Na marcação completa,
 * os mesmos — a revisão sai sublinhada e riscada, na cor do autor, como na tela.
 * Na simples e na sem marcação, o texto final: o excluído sai, o inserido fica
 * como texto comum. No Original, o contrário. A marca de parágrafo e a linha de
 * tabela seguem a mesma regra, e o bloco que a tela esconde inteiro não vai.
 */
export function blocksForView(blocks: readonly ProseMirrorNode[], view: RevisionView): ProseMirrorNode[] {
  if (view === RevisionView.All) return [...blocks]
  return blocks.flatMap((block) => nodeForView(block, view) ?? [])
}

/** A referência de nota sem as marcas de revisão → a do documento, que tem o rótulo. */
const originalNoteRefs = new WeakMap<ProseMirrorNode, ProseMirrorNode>()

/**
 * Os rótulos das referências que a folha mostra, na ordem: o que o modo de
 * mostrar escondeu não está no HTML, e não leva rótulo.
 */
function printedNoteLabels(
  fragments: readonly ProseMirrorNode[],
  labelOf: ReadonlyMap<ProseMirrorNode, string>,
): string[] {
  const labels: string[] = []
  for (const fragment of fragments) {
    fragment.descendants((node) => {
      if (node.type.name !== 'noteRef') return true
      labels.push(labelOf.get(originalNoteRefs.get(node) ?? node) ?? '')
      return false
    })
  }
  return labels
}

function nodeForView(node: ProseMirrorNode, view: RevisionView): ProseMirrorNode | null {
  if (node.isInline) {
    if (isHiddenInline(node, view)) return null
    const marks = node.marks.filter((mark) => mark.type.name !== INSERTION && mark.type.name !== DELETION)
    if (marks.length === node.marks.length) return node
    const shown = node.mark(marks)
    if (node.type.name === 'noteRef') originalNoteRefs.set(shown, node)
    return shown
  }
  if (isHiddenBlock(node, view)) return null

  const children: ProseMirrorNode[] = []
  node.forEach((child) => {
    const shown = nodeForView(child, view)
    if (shown !== null) children.push(shown)
  })
  const attrs = { ...node.attrs }
  if ('markRevision' in attrs) attrs['markRevision'] = null
  if ('rowRevision' in attrs) attrs['rowRevision'] = null
  if (children.length === 0 && node.childCount > 0 && !node.isTextblock) {
    // A célula vazia continua (a linha precisa dela); a tabela ou a lista que
    // perdeu tudo, não.
    return node.type.name === 'tableCell' || node.type.name === 'tableHeader'
      ? node.type.createAndFill(attrs, null, node.marks)
      : null
  }
  return node.type.create(attrs, Fragment.fromArray(children), node.marks)
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
 * Justificada pelo mesmo truque do espaçador da tela: um elemento da largura da
 * linha no fim, que só cabe numa linha própria e faz da anterior uma quebra
 * automática. O `text-align-last` parecia equivalente e não era: ele vale também
 * para a linha antes de cada `<br>` (Shift+Enter), que a tela deixa à esquerda.
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
    if (justified) {
      const filler = document.createElement('span')
      filler.setAttribute('aria-hidden', 'true')
      filler.style.cssText = 'display:inline-block;width:100%;height:0;vertical-align:top'
      last.appendChild(filler)
    }
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
