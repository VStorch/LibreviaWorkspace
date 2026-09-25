import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { paginate, type MeasuredBlock } from '@services/document/paginate.js'
import { effectiveAttrs } from '@services/document/style-cascade.js'
import type { StyleSheet } from '@services/document/styles.js'
import {
  contentHeightMm,
  contentInsetsMm,
  mmToPx,
  pageDimensionsMm,
  type PageSetup,
} from '@services/document/model.js'
import { NO_BANDS, type BandHeights } from '@services/document/band.js'
import { applyPageGaps, type RepeatedHeader } from './extensions/pagination.js'
import { LINE_GAP_CLASS, measureLines } from './line-boxes.js'

/** Espaço entre uma folha e a seguinte, como numa pilha de papel. */
export const SHEET_GUTTER_PX = 28

/** Nenhum vao aplicado. Constante para `sameGaps` poder compara-la por valor. */
const EMPTY_GAPS = new Map<number, number>()

export interface PageLayout {
  /** Quantas folhas desenhar. */
  readonly pages: number
  /** Altura total da pilha, com os vãos. */
  readonly stackHeightPx: number
  /** Topo de cada folha, em pixels, dentro da pilha. */
  readonly sheetTops: readonly number[]
  readonly sheetHeights: readonly number[]
  /**
   * Bloco e eventual linha/item que abrem cada folha a partir da segunda.
   *
   * É o que permite ao papel sair das mesmas páginas que a tela: recortar a
   * lista de blocos nestes pontos dá as folhas prontas, sem ninguém repaginar.
   */
  readonly pageStarts: readonly PageStart[]
  /**
   * Em que folha cada bloco caiu, e a que altura dentro dela.
   *
   * É o que falta para desenhar um objeto ancorado: a posição dele vem em
   * relação ao parágrafo âncora, e o parágrafo só tem posição depois de paginar.
   */
  readonly anchors: readonly BlockAnchor[]
}

export interface PageStart {
  readonly blockIndex: number
  /** Índice da linha ou item que abre a folha, quando o corte é interno. */
  readonly childIndex?: number
  /**
   * Onde o parágrafo recomeça, quando a folha o corta entre linhas: posição
   * dentro do conteúdo dele, a mesma que `Node.cut` recebe.
   */
  readonly offset?: number
  /** A folha abre com as linhas de cabeçalho da tabela repetidas. */
  readonly repeatHeader?: boolean
}

/** O corte cai dentro do bloco — o bloco começa na folha anterior. */
export function isInternalStart(start: PageStart): boolean {
  return start.childIndex !== undefined || start.offset !== undefined
}

interface CutTarget {
  readonly at: number
  readonly start: PageStart
  readonly nodes: readonly { position: number; natural: number }[]
  /**
   * Corte entre linhas de um parágrafo: a posição do primeiro caractere da
   * linha, resolvida só se o corte for escolhido, e a do próprio parágrafo.
   */
  readonly line?: { readonly resolve: () => number | null; readonly block: number }
  /** Corte entre linhas de tabela com cabeçalho: o que se repete no alto da folha. */
  readonly header?: () => RepeatedHeader
}

export interface BlockAnchor {
  readonly pageIndex: number
  /** Topo do bloco dentro da folha, em pixels, já incluída a margem superior. */
  readonly topPx: number
}

/**
 * Mede o documento, decide onde as páginas quebram e empurra os blocos.
 *
 * O laço fecha sozinho porque a medição é convertida para **coordenadas de
 * fluxo** antes de decidir: o `offsetTop` que o navegador dá já inclui os vãos
 * que aplicamos na passada anterior, então subtraí-los devolve a altura que o
 * documento teria como tira contínua. Decidir sobre essa altura é estável —
 * aplicar o resultado não muda a entrada da próxima medida. Sem isso, cada
 * passada empurraria os blocos um pouco mais e a paginação nunca assentaria.
 */
export function usePagination(
  editor: Editor | null,
  page: PageSetup,
  revision: number,
  bands: BandHeights = NO_BANDS,
  /**
   * Se os vãos devem ser **empurrados no DOM**.
   *
   * O modo de leitura desliga isto, e só isto: a conta continua acontecendo, e
   * `pageStarts` continua valendo. É de propósito, e é o que permite imprimir
   * de dentro do modo de leitura sem sair dele — o papel sai com as mesmas
   * folhas de sempre, porque as coordenadas de fluxo não dependem de os vãos
   * estarem aplicados. Elas são, por definição, a altura que o documento teria
   * como tira contínua, que é exatamente o que o modo de leitura mostra.
   *
   * Desligar a medição junto pareceria mais simples e custaria a impressão: o
   * gravador lê `layout` no momento de imprimir, e um layout de uma página só
   * mandaria o documento inteiro para uma folha.
   */
  paginated = true,
  /**
   * Os estilos do documento: o "manter com o próximo" pode vir do estilo, e o
   * bloco só carrega o que o parágrafo declara.
   */
  styles: StyleSheet | null = null,
): PageLayout {
  const [layout, setLayout] = useState<PageLayout>({
    pages: 1,
    stackHeightPx: 0,
    sheetTops: [0],
    sheetHeights: [],
    pageStarts: [],
    anchors: [],
  })

  /**
   * Vãos que já estão aplicados no DOM.
   *
   * Mora numa `ref`, e não numa variável do efeito, porque o efeito é refeito a
   * cada tecla: esquecer o que já foi empurrado faria a leitura seguinte tomar
   * o `offsetTop` empurrado como se fosse altura de fluxo. As folhas mudavam de
   * quantidade a cada letra digitada, e a quebra pedida à mão chegava a
   * desaparecer.
   *
   * A chave passou a ser a **posição** do nó no documento, e não o índice do
   * bloco: um corte interno empurra uma linha de tabela ou um item de lista, e
   * nenhum dos dois tem índice na lista de blocos de primeiro nível.
   */
  const applied = useRef(new Map<number, number>())

  /**
   * O que a decoração recebeu, que é o vão **mais** a margem natural.
   *
   * Dois mapas, e não um, pela mesma razão de sempre: o vão é o que a conta de
   * fluxo desconta, e o valor escrito é o que o CSS lê. Comparar o escrito é o
   * que evita uma transação idêntica a cada medição; descontá-lo faria a conta
   * errar por uma margem natural a cada corte.
   */
  const lastWritten = useRef(new Map<number, number>())

  /** Os vãos entre linhas de parágrafo cortado, por posição do espaçador. */
  const lastLines = useRef(new Map<number, number>())

  /** Os cabeçalhos de tabela repetidos, comparados pelo que desenham. */
  const lastHeaders = useRef('[]')

  useEffect(() => {
    if (editor === null) return undefined

    const element = editor.view.dom as HTMLElement // alvo do observador de tamanho
    const pageHeightPx = mmToPx(pageDimensionsMm(page).height)
    const contentHeightPx = mmToPx(contentHeightMm(page, bands))
    // A margem é um piso: um cabeçalho mais alto que ela empurra o corpo para
    // baixo, e é essa a altura de onde a folha seguinte recomeça.
    const insets = contentInsetsMm(page, bands)
    const marginTopPx = mmToPx(insets.top)
    const marginBottomPx = mmToPx(insets.bottom)

    const measure = (): void => {
      // Os espaçadores entre linhas saem de cena durante a medida. Diferente do
      // vão de bloco, eles mudam **onde as linhas quebram**: o espaçador ocupa
      // uma linha inteira, e quando o texto acima encolhe e ele deixa de estar
      // num começo de linha, força uma quebra ali — a medida seguinte achava o
      // mesmo corte, e a folha ficava com linhas vazias no pé para sempre.
      // Escondidos, as linhas são as do texto; o estilo volta no mesmo quadro,
      // antes de o navegador desenhar, e o observador de tamanho não vê nada.
      const lineGapsInDom = Array.from(element.querySelectorAll<HTMLElement>(`.${LINE_GAP_CLASS}`))
      for (const gap of lineGapsInDom) gap.style.display = 'none'
      try {
        measureHidden()
      } finally {
        for (const gap of lineGapsInDom) gap.style.display = 'inline-block'
      }
    }

    const measureHidden = (): void => {
      // Percorrido pelo **documento**, e não pelos filhos do DOM: os dois não
      // são o mesmo sistema de índices. Um documento do corpus tem 15 elementos
      // na tela e 17 nós no topo do modelo, e a diferença é silenciosa — a
      // decoração cairia num bloco e o recorte do papel noutro, cada um errando
      // por uma quantidade diferente. `nodeDOM` liga um ao outro.
      let accumulated = 0
      const blocks: MeasuredBlock[] = []
      const targets: CutTarget[] = []
      const origin = offsetTopOf(element)

      editor.state.doc.forEach((block, offset, blockIndex) => {
        const dom = editor.view.nodeDOM(offset)
        const node = dom instanceof HTMLElement ? dom : null
        if (node === null) {
          blocks.push({
            top: 0,
            height: 0,
            breakpoints: [],
            isPageBreak: false,
            breakAfter: false,
            keepWithNext: false,
          })
          return
        }

        accumulated += shiftOf(node)
        const top = offsetTopOf(node) - origin - accumulated
        const before = blocks.at(-1)
        targets.push({
          at: top,
          start: { blockIndex },
          nodes: [
            {
              position: offset,
              natural: Math.max(top - (before === undefined ? 0 : before.top + before.height), 0),
            },
          ],
        })

        // O TableView redimensionável envolve a tabela num div. Só as linhas
        // da tabela externa contam; tabelas aninhadas pertencem às células.
        const table =
          node instanceof HTMLTableElement ? node : node.querySelector<HTMLTableElement>(':scope > table')
        const children =
          table !== null
            ? Array.from(table.rows)
            : node.tagName === 'UL' || node.tagName === 'OL'
              ? Array.from(node.children).filter(
                  (child): child is HTMLElement => child instanceof HTMLElement && child.tagName === 'LI',
                )
              : []
        let internal = 0
        const breakpoints: number[] = []

        // Parágrafo e título cortam entre linhas. As linhas medem a partir da
        // borda do bloco, e o topo de fluxo dele já está em `top`.
        const lines = block.isTextblock && children.length === 0 ? measureLines(editor.view, node) : null
        if (lines !== null) {
          lines.starts.forEach((start, index) => {
            const at = top + start
            breakpoints.push(at)
            targets.push({
              at,
              start: { blockIndex },
              nodes: [],
              line: { resolve: () => lines.positionOf(index), block: offset },
            })
          })
          internal = lines.shift
        }

        // A captura ancorada: o parágrafo dela tem uma linha vazia depois do
        // quadro (o `::after` de 1lh em `content-styles.ts`), e o LibreOffice a
        // deixa passar para a folha seguinte quando ela não cabe — o quadro
        // fica. O corte é no pé do quadro, e o espaçador entra no fim do
        // parágrafo, depois da imagem.
        const freeBreakpoints: number[] = []
        let hangingBottom = 0
        if (
          lines !== null &&
          node.querySelector(':scope > .node-image[data-anchored], :scope > img[data-anchored]') !== null
        ) {
          // Com texto, a linha do parágrafo é a do texto, logo abaixo do quadro
          // quando ele abre o parágrafo: o corte entre o quadro e ela é livre
          // da regra de viúvas, como o da linha vazia.
          const first = lines.starts[0]
          if (
            first !== undefined &&
            block.firstChild?.type.name === 'image' &&
            block.firstChild.attrs['anchored'] === true
          ) {
            freeBreakpoints.push(top + first)
          }
          // Sem texto, a linha vazia de 1lh depois do quadro pode sobrar no pé
          // da folha: o LibreOffice a deixa passar da margem de baixo, e o bloco
          // seguinte abre a folha nova sem ela.
          const after = parseFloat(getComputedStyle(node, '::after').height)
          if (Number.isFinite(after) && after > 0) hangingBottom = after
        }

        // Linhas de cabeçalho (`w:tblHeader`, células `th`) no começo da tabela:
        // repetem-se no alto de cada folha em que a tabela continua. Cortar
        // dentro delas, ou logo depois, deixaria o cabeçalho sozinho no pé.
        const rows = table !== null ? Array.from(table.rows) : []
        let headerRows = 0
        while (
          headerRows < rows.length - 1 &&
          rows[headerRows]!.cells.length > 0 &&
          Array.from(rows[headerRows]!.cells).every((cell) => cell.tagName === 'TH')
        ) {
          headerRows += 1
        }
        const lastHeader = rows[headerRows - 1]
        const repeatHeight =
          lastHeader === undefined
            ? 0
            : offsetTopOf(lastHeader) + lastHeader.offsetHeight - offsetTopOf(rows[0]!)

        children.forEach((child, childIndex) => {
          const cells = child instanceof HTMLTableRowElement ? Array.from(child.cells) : []
          const shift = cells.length > 0 ? shiftOf(cells[0]!) : shiftOf(child)
          // Padding aumenta a linha para baixo; margem já deslocou seu topo.
          const at = offsetTopOf(child) - origin - accumulated - internal - (cells.length === 0 ? shift : 0)
          if (childIndex > headerRows) {
            breakpoints.push(at)
            targets.push({
              at,
              start: { blockIndex, childIndex },
              nodes: (cells.length > 0 ? cells : [child]).map((target) => ({
                position: editor.view.posAtDOM(target, 0) - 1,
                natural:
                  parseFloat(
                    cells.length > 0
                      ? getComputedStyle(target).paddingTop
                      : getComputedStyle(target).marginTop,
                  ) - shiftOf(target),
              })),
              ...(table !== null && repeatHeight > 0
                ? {
                    header: () =>
                      repeatedHeader(editor, table, rows.slice(0, headerRows), cells[0]!, repeatHeight),
                  }
                : {}),
            })
          }
          internal += shift
        })
        const effective = effectiveAttrs(block, styles)
        blocks.push({
          top,
          height: node.offsetHeight - internal,
          breakpoints,
          isPageBreak: node.hasAttribute('data-page-break'),
          breakAfter: node.hasAttribute('data-break-after'),
          keepWithNext: effective['keepNext'] === true || /^H[1-6]$/.test(node.tagName),
          keepLines: effective['keepLines'] === true,
          widowControl: lines !== null && effective['widowControl'] !== false,
          ...(repeatHeight > 0 ? { repeatHeight } : {}),
          ...(freeBreakpoints.length > 0 ? { freeBreakpoints } : {}),
          ...(hangingBottom > 0 ? { hangingBottom } : {}),
        })
        accumulated += internal
      })

      const breaks = paginate(blocks, contentHeightPx)

      // Vão = o que sobrou da folha + as duas margens + o espaço entre papéis.
      // É essa soma que faz o bloco cair exatamente no topo da coluna de texto
      // da folha seguinte.
      // Dois números por bloco, e não um: o **empurrão** e a **margem escrita**.
      //
      // O estilo da decoração é acrescentado ao do nó, e em CSS a última
      // declaração ganha — então a margem do vão não se soma à margem natural
      // do bloco, ela a substitui. Escrever só o empurrão faria o bloco subir o
      // tanto da margem que ele já tinha, e a conta de fluxo, que desconta o
      // empurrão, passaria a errar por essa diferença. Num título com 18 pt de
      // espaço antes, isso bastava para o corte cair um bloco adiante — a tela
      // mostrava o título abrindo a folha e o papel o deixava no fim da
      // anterior.
      //
      // A margem natural é observável mesmo depois de decorada: as coordenadas
      // de fluxo já removem o empurrão, então a distância entre o fim de um
      // bloco e o começo do seguinte é a margem que o documento pede.
      const gaps = new Map<number, number>()
      const written = new Map<number, number>()
      const lineGaps = new Map<number, number>()
      const headers: RepeatedHeader[] = []
      let previous = 0
      const pageStarts: PageStart[] = []
      const sheetHeights: number[] = []

      // Índice dos cortes internos por altura, e um cursor para os de bloco: a
      // lista sai da medida em ordem de fluxo, e os cortes também crescem, então
      // cada folha custa uma consulta, e não uma varredura do documento.
      const internalAt = new Map<number, CutTarget>()
      for (const target of targets) {
        if (
          (target.start.childIndex !== undefined || target.line !== undefined) &&
          !internalAt.has(target.at)
        ) {
          internalAt.set(target.at, target)
        }
      }
      let cursor = 0
      const blockTargetFrom = (at: number): CutTarget | undefined => {
        while (cursor < targets.length && (targets[cursor]!.at < at || targets[cursor]!.line !== undefined))
          cursor++
        return targets[cursor]
      }

      for (const at of breaks) {
        const internal = internalAt.get(at)
        const position = internal?.line?.resolve() ?? null
        // A linha cujo caractere não se achou (DOM trocado no meio da medida)
        // cede ao bloco seguinte: pior a folha curta que um espaçador perdido.
        const target =
          internal !== undefined && (internal.line === undefined || position !== null)
            ? internal
            : blockTargetFrom(at)
        // A linha vazia da captura que sobra no pé (`hangingBottom`) cabe na
        // margem de baixo: nem estica a folha, nem empurra o bloco seguinte.
        const span = at - previous
        const hung = Math.min(Math.max(span - contentHeightPx, 0), hangingAt(blocks, at))
        const used = span - hung
        sheetHeights.push(Math.max(pageHeightPx, used + marginTopPx + marginBottomPx))
        const shift =
          Math.max(contentHeightPx - used, 0) - hung + marginBottomPx + SHEET_GUTTER_PX + marginTopPx
        if (target?.line !== undefined && position !== null) {
          // O espaçador entra antes do primeiro caractere da linha; o papel
          // recorta o parágrafo no mesmo caractere.
          pageStarts.push({ ...target.start, offset: position - target.line.block - 1 })
          lineGaps.set(position, shift)
        } else if (target !== undefined) {
          // O cabeçalho repetido mora no vão, entre o topo da folha e a linha:
          // o vão cresce a altura dele, e a conta de fluxo desconta os dois.
          const header = target.header?.()
          const extra = header !== undefined && header.height < contentHeightPx / 2 ? header.height : 0
          if (header !== undefined && extra > 0) headers.push(header)
          pageStarts.push(extra > 0 ? { ...target.start, repeatHeader: true } : target.start)
          for (const node of target.nodes) {
            gaps.set(node.position, shift + extra)
            written.set(node.position, shift + extra + node.natural)
          }
          // A folha nova começa acima do corte, pela altura do cabeçalho.
          previous = at - extra
          continue
        } else {
          // Uma quebra explícita final ainda abre uma folha vazia.
          pageStarts.push({ blockIndex: blocks.length })
        }
        previous = at
      }

      const bottom = blocks.reduce((bottom, block) => Math.max(bottom, block.top + block.height), 0)
      const lastSpan = bottom - previous
      const lastHung = Math.min(Math.max(lastSpan - contentHeightPx, 0), hangingAt(blocks, bottom))
      sheetHeights.push(Math.max(pageHeightPx, lastSpan - lastHung + marginTopPx + marginBottomPx))
      const sheetTops: number[] = []
      let stackHeightPx = 0
      for (const height of sheetHeights) {
        sheetTops.push(stackHeightPx)
        stackHeightPx += height + SHEET_GUTTER_PX
      }

      // No modo de leitura o documento é uma tira contínua: os vãos saem do
      // DOM, e o mapa do que está aplicado esvazia junto. Esvaziá-lo é o que
      // importa — a medição seguinte desconta o que este mapa diz estar
      // empurrado, e deixá-lo cheio faria toda altura ser lida a menos.
      const target = paginated ? gaps : EMPTY_GAPS
      const targetWritten = paginated ? written : EMPTY_GAPS
      const targetLines = paginated ? lineGaps : EMPTY_GAPS
      const targetHeaders = paginated ? headers : []
      const headersKey = JSON.stringify(targetHeaders)

      if (
        !sameGaps(applied.current, target) ||
        !sameGaps(lastWritten.current, targetWritten) ||
        !sameGaps(lastLines.current, targetLines) ||
        lastHeaders.current !== headersKey
      ) {
        lastHeaders.current = headersKey
        applied.current = target
        lastWritten.current = targetWritten
        lastLines.current = targetLines
        applyPageGaps(editor.view, targetWritten, target, targetLines, targetHeaders)
      }

      setLayout({
        pages: sheetHeights.length,
        stackHeightPx: stackHeightPx - SHEET_GUTTER_PX,
        sheetTops,
        sheetHeights,
        pageStarts,
        anchors: anchorsFor(blocks, breaks, marginTopPx),
      })
    }

    // Uma medida por quadro, no máximo. Digitar depressa dispara dezenas de
    // atualizações por segundo, e medir em todas custa layout do navegador sem
    // mudar resposta nenhuma — a folha não nasce entre duas teclas.
    let scheduled = 0
    const schedule = (): void => {
      if (scheduled !== 0) return
      scheduled = requestAnimationFrame(() => {
        scheduled = 0
        measure()
      })
    }

    schedule()

    // Altura muda ao digitar, ao carregar imagem e ao trocar a fonte.
    const observer = new ResizeObserver(schedule)
    observer.observe(element)
    return () => {
      observer.disconnect()
      if (scheduled !== 0) cancelAnimationFrame(scheduled)
    }
  }, [editor, page, revision, bands.headerMm, bands.footerMm, paginated, styles])

  return layout
}

/**
 * Quanto do fim da folha que termina em `at` pode passar do pé: a linha vazia
 * da captura (`hangingBottom`) e o vão até o bloco que abre a folha seguinte.
 */
function hangingAt(blocks: readonly MeasuredBlock[], at: number): number {
  const ending = blocks.filter((block) => block.height > 0 && block.top + block.height <= at + 0.5).at(-1)
  if (ending === undefined || (ending.hangingBottom ?? 0) <= 0) return 0
  return ending.hangingBottom! + Math.max(at - (ending.top + ending.height), 0)
}

function sameGaps(left: ReadonlyMap<number, number>, right: ReadonlyMap<number, number>): boolean {
  if (left.size !== right.size) return false
  for (const [index, gap] of left) {
    // Um pixel de diferença não vale uma nova transação: o arredondamento da
    // medida oscila sozinho, e redesenhar a cada oscilação faria o documento
    // tremer enquanto se digita.
    if (!right.has(index) || Math.abs(right.get(index)! - gap) > 0.5) return false
  }
  return true
}

/**
 * Em que folha cada bloco caiu, e a que altura dentro dela.
 *
 * Em coordenadas de fluxo os cortes são fronteiras crescentes, então uma
 * varredura só resolve — os blocos já vêm em ordem.
 */
function anchorsFor(
  blocks: readonly MeasuredBlock[],
  breaks: readonly number[],
  marginTopPx: number,
): BlockAnchor[] {
  const anchors: BlockAnchor[] = []
  let page = 0

  for (const block of blocks) {
    while (page < breaks.length && block.top >= breaks[page]!) page += 1
    const start = page === 0 ? 0 : breaks[page - 1]!
    anchors.push({ pageIndex: page, topPx: marginTopPx + (block.top - start) })
  }

  return anchors
}

/** Soma as origens dos offsetParents: uma linha mede a partir da tabela. */
function offsetTopOf(node: HTMLElement): number {
  let top = 0
  let current: HTMLElement | null = node
  while (current !== null) {
    top += current.offsetTop
    current = current.offsetParent instanceof HTMLElement ? current.offsetParent : null
  }
  return top
}

/** A decoração acompanha edições no modelo; a medida lê o empurrão já mapeado. */
function shiftOf(node: HTMLElement): number {
  return Number(node.dataset.pageShift ?? 0)
}

/**
 * O cabeçalho que se repete no alto da folha, pronto para a decoração.
 *
 * Uma cópia das linhas de cabeçalho, numa tabela com as mesmas colunas e a
 * mesma largura, posta sobre o vão da linha que abre a folha. Mora na primeira
 * célula dessa linha e sai dela por margens negativas: o ponto de partida é o
 * canto do conteúdo da célula, que é onde o elemento fora do fluxo começaria.
 */
function repeatedHeader(
  editor: Editor,
  table: HTMLTableElement,
  headerRows: readonly HTMLTableRowElement[],
  cell: HTMLTableCellElement,
  height: number,
): RepeatedHeader {
  const tableBox = table.getBoundingClientRect()
  const cellBox = cell.getBoundingClientRect()
  const scale = table.offsetWidth > 0 && tableBox.width > 0 ? tableBox.width / table.offsetWidth : 1
  const style = getComputedStyle(cell)
  const left = (cellBox.left - tableBox.left) / scale + cell.clientLeft + parseFloat(style.paddingLeft)
  const natural = parseFloat(style.paddingTop) - shiftOf(cell)
  const colgroup = table.querySelector(':scope > colgroup')?.outerHTML ?? ''
  const html =
    `<table class="${table.className}" style="width:${table.offsetWidth}px;margin:0">${colgroup}` +
    `<tbody>${headerRows.map(cleanHeaderRow).join('')}</tbody></table>`
  return {
    position: editor.view.posAtDOM(cell, 0),
    html,
    height,
    offsetTop: height + Math.max(natural, 0),
    offsetLeft: left,
  }
}

/**
 * A cópia do cabeçalho sem o que é do momento, e não do documento: a célula
 * selecionada, o destaque da busca, a alça de arrastar coluna, o vão de página
 * e o espaçador de linha. Copiados, eles apareciam repetidos em cada folha.
 */
function cleanHeaderRow(row: HTMLTableRowElement): string {
  const copy = row.cloneNode(true) as HTMLTableRowElement
  for (const transient of copy.querySelectorAll(
    '.column-resize-handle, .page-line-gap, .page-repeated-header',
  )) {
    transient.remove()
  }
  for (const element of [copy, ...copy.querySelectorAll<HTMLElement>('*')]) {
    element.classList.remove('selectedCell', 'search-hit', 'search-hit--current', 'ProseMirror-selectednode')
    if (element.hasAttribute('data-page-start')) {
      element.style.removeProperty('padding-top')
      element.style.removeProperty('margin-top')
      element.removeAttribute('data-page-start')
      element.removeAttribute('data-page-shift')
    }
  }
  return copy.outerHTML
}
