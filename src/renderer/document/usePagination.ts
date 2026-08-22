import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { paginate, type MeasuredBlock } from '@services/document/paginate.js'
import { contentHeightMm, mmToPx, pageDimensionsMm, type PageSetup } from '@services/document/model.js'
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
export function usePagination(editor: Editor | null, page: PageSetup, revision: number): PageLayout {
  const [layout, setLayout] = useState<PageLayout>({ pages: 1, stackHeightPx: 0, sheetTops: [0] })

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

    const element = editor.view.dom as HTMLElement
    const pageHeightPx = mmToPx(pageDimensionsMm(page).height)
    const contentHeightPx = mmToPx(contentHeightMm(page))
    const marginTopPx = mmToPx(page.margins.top)
    const marginBottomPx = mmToPx(page.margins.bottom)

    const measure = (): void => {
      const children = Array.from(element.children) as HTMLElement[]

      let accumulated = 0
      const blocks: MeasuredBlock[] = children.map((node, index) => {
        accumulated += applied.current.get(index) ?? 0
        return {
          top: node.offsetTop - accumulated,
          height: node.offsetHeight,
          isPageBreak: node.hasAttribute('data-page-break'),
          keepWithNext: node.hasAttribute('data-keep-next') || /^H[1-6]$/.test(node.tagName),
        }
      })

      const breaks = paginate(blocks, contentHeightPx)

      // Vão = o que sobrou da folha + as duas margens + o espaço entre papéis.
      // É essa soma que faz o bloco cair exatamente no topo da coluna de texto
      // da folha seguinte.
      const gaps = new Map<number, number>()
      let previous = 0
      for (const at of breaks) {
        const index = blocks.findIndex((block) => block.top >= at)
        if (index <= 0) continue

        const remaining = Math.max(contentHeightPx - (at - previous), 0)
        gaps.set(index, remaining + marginBottomPx + SHEET_GUTTER_PX + marginTopPx)
        previous = at
      }

      if (!sameGaps(applied.current, gaps)) {
        applied.current = gaps
        applyPageGaps(editor.view, gaps)
      }

      const pages = breaks.length + 1
      setLayout({
        pages,
        stackHeightPx: pages * pageHeightPx + (pages - 1) * SHEET_GUTTER_PX,
        sheetTops: Array.from({ length: pages }, (_, i) => i * (pageHeightPx + SHEET_GUTTER_PX)),
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
  }, [editor, page, revision])

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
