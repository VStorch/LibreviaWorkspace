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
import { applyPageGaps } from './extensions/pagination.js'
import { measureLines } from './line-boxes.js'

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

        children.forEach((child, childIndex) => {
          const cells = child instanceof HTMLTableRowElement ? Array.from(child.cells) : []
          const shift = cells.length > 0 ? shiftOf(cells[0]!) : shiftOf(child)
          // Padding aumenta a linha para baixo; margem já deslocou seu topo.
          const at = offsetTopOf(child) - origin - accumulated - internal - (cells.length === 0 ? shift : 0)
          if (childIndex > 0) {
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
      let previous = 0
      const pageStarts: PageStart[] = []
      const sheetHeights: number[] = []

      for (const at of breaks) {
        const internal = targets.find(
          (target) =>
            target.at === at && (target.start.childIndex !== undefined || target.line !== undefined),
        )
        const position = internal?.line?.resolve() ?? null
        // A linha cujo caractere não se achou (DOM trocado no meio da medida)
        // cede ao bloco seguinte: pior a folha curta que um espaçador perdido.
        const target =
          internal !== undefined && (internal.line === undefined || position !== null)
            ? internal
            : targets.find((target) => target.at >= at && target.line === undefined)
        const used = at - previous
        sheetHeights.push(Math.max(pageHeightPx, used + marginTopPx + marginBottomPx))
        const shift = Math.max(contentHeightPx - used, 0) + marginBottomPx + SHEET_GUTTER_PX + marginTopPx
        if (target?.line !== undefined && position !== null) {
          // O espaçador entra antes do primeiro caractere da linha; o papel
          // recorta o parágrafo no mesmo caractere.
          pageStarts.push({ ...target.start, offset: position - target.line.block - 1 })
          lineGaps.set(position, shift)
        } else if (target !== undefined) {
          pageStarts.push(target.start)
          for (const node of target.nodes) {
            gaps.set(node.position, shift)
            written.set(node.position, shift + node.natural)
          }
        } else {
          // Uma quebra explícita final ainda abre uma folha vazia.
          pageStarts.push({ blockIndex: blocks.length })
        }
        previous = at
      }

      const bottom = blocks.reduce((bottom, block) => Math.max(bottom, block.top + block.height), 0)
      sheetHeights.push(Math.max(pageHeightPx, bottom - previous + marginTopPx + marginBottomPx))
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

      if (
        !sameGaps(applied.current, target) ||
        !sameGaps(lastWritten.current, targetWritten) ||
        !sameGaps(lastLines.current, targetLines)
      ) {
        applied.current = target
        lastWritten.current = targetWritten
        lastLines.current = targetLines
        applyPageGaps(editor.view, targetWritten, target, targetLines)
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
