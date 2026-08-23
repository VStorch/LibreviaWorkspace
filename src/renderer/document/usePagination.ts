import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { paginate, type MeasuredBlock } from '@services/document/paginate.js'
import {
  contentHeightMm,
  contentInsetsMm,
  mmToPx,
  NO_BANDS,
  pageDimensionsMm,
  type BandHeights,
  type PageSetup,
} from '@services/document/model.js'
import { applyPageGaps } from './extensions/pagination.js'

/** Espaço entre uma folha e a seguinte, como numa pilha de papel. */
export const SHEET_GUTTER_PX = 28

export interface PageLayout {
  /** Quantas folhas desenhar. */
  readonly pages: number
  /** Altura total da pilha, com os vãos. */
  readonly stackHeightPx: number
  /** Topo de cada folha, em pixels, dentro da pilha. */
  readonly sheetTops: readonly number[]
  /**
   * Índice do bloco que abre cada folha a partir da segunda.
   *
   * É o que permite ao papel sair das mesmas páginas que a tela: recortar a
   * lista de blocos nestes pontos dá as folhas prontas, sem ninguém repaginar.
   */
  readonly pageStarts: readonly number[]
  /**
   * Em que folha cada bloco caiu, e a que altura dentro dela.
   *
   * É o que falta para desenhar um objeto ancorado: a posição dele vem em
   * relação ao parágrafo âncora, e o parágrafo só tem posição depois de paginar.
   */
  readonly anchors: readonly BlockAnchor[]
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
): PageLayout {
  const [layout, setLayout] = useState<PageLayout>({
    pages: 1,
    stackHeightPx: 0,
    sheetTops: [0],
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
   */
  const applied = useRef(new Map<number, number>())

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
      let index = 0

      editor.state.doc.forEach((_node, offset) => {
        const dom = editor.view.nodeDOM(offset)
        const node = dom instanceof HTMLElement ? dom : null
        accumulated += applied.current.get(index) ?? 0
        index += 1

        blocks.push(
          node === null
            ? { top: 0, height: 0, isPageBreak: false, breakAfter: false, keepWithNext: false }
            : {
                top: node.offsetTop - accumulated,
                height: node.offsetHeight,
                isPageBreak: node.hasAttribute('data-page-break'),
                breakAfter: node.hasAttribute('data-break-after'),
                keepWithNext: node.hasAttribute('data-keep-next') || /^H[1-6]$/.test(node.tagName),
              },
        )
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
      let previous = 0

      for (const at of breaks) {
        const index = blocks.findIndex((block) => block.top >= at)
        if (index <= 0) continue

        const block = blocks[index]!
        const before = blocks[index - 1]!
        const natural = Math.max(block.top - (before.top + before.height), 0)

        const remaining = Math.max(contentHeightPx - (at - previous), 0)
        const shift = remaining + marginBottomPx + SHEET_GUTTER_PX + marginTopPx

        gaps.set(index, shift)
        written.set(index, shift + natural)
        previous = at
      }

      if (!sameGaps(applied.current, gaps)) {
        applied.current = gaps
        applyPageGaps(editor.view, written)
      }

      const pages = breaks.length + 1
      setLayout({
        pages,
        stackHeightPx: pages * pageHeightPx + (pages - 1) * SHEET_GUTTER_PX,
        sheetTops: Array.from({ length: pages }, (_, i) => i * (pageHeightPx + SHEET_GUTTER_PX)),
        pageStarts: [...gaps.keys()].sort((a, b) => a - b),
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
  }, [editor, page, revision, bands.headerMm, bands.footerMm])

  return layout
}

function sameGaps(left: ReadonlyMap<number, number>, right: ReadonlyMap<number, number>): boolean {
  if (left.size !== right.size) return false
  for (const [index, gap] of left) {
    // Um pixel de diferença não vale uma nova transação: o arredondamento da
    // medida oscila sozinho, e redesenhar a cada oscilação faria o documento
    // tremer enquanto se digita.
    if (Math.abs((right.get(index) ?? Number.NaN) - gap) > 1) return false
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
