import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import { LINE_GAP_CLASS } from '../line-boxes.js'

/**
 * O bloco que abre uma folha ganha uma margem do tamanho do que sobrou da
 * anterior, mais as margens e o vão entre papéis. Como **decoração**, e não
 * elemento: um espaçador de verdade entraria na seleção, no `Ctrl+A` e no que se
 * copia.
 */
export const paginationKey = new PluginKey<DecorationSet>('pagination')

export const Pagination = Extension.create({
  name: 'pagination',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: paginationKey,
        state: {
          init: () => DecorationSet.empty,
          apply(transaction, current) {
            const next = transaction.getMeta(paginationKey) as DecorationSet | undefined
            if (next !== undefined) return next

            // Até a medição chegar, as decorações acompanham a edição, senão as folhas piscariam.
            return current.map(transaction.mapping, transaction.doc)
          },
        },
        props: {
          decorations: (state) => paginationKey.getState(state),
        },
      }),
      // A captura ancorada com texto: a linha de 1lh (`content-styles.ts`) não se soma.
      new Plugin({
        props: {
          decorations: (state) => {
            const decorations: Decoration[] = []
            state.doc.forEach((node, offset) => {
              if (hasAnchoredImageAndText(node)) {
                decorations.push(Decoration.node(offset, offset + node.nodeSize, { [ANCHOR_TEXT_ATTR]: '' }))
              }
            })
            return DecorationSet.create(state.doc, decorations)
          },
        },
      }),
    ]
  },
})

/** Fora do histórico: desfazer volta o que a pessoa escreveu, e não onde a página caiu. */
export function applyPageGaps(
  view: EditorView,
  written: ReadonlyMap<number, number>,
  gaps: ReadonlyMap<number, number>,
  { lines = new Map(), headers = [], columns = new Map() }: PageGapExtras = {},
): void {
  const decorations: Decoration[] = []

  // O cabeçalho repetido é cópia num widget: fora da seleção e da edição.
  for (const header of headers) {
    if (header.position <= 0 || header.position > view.state.doc.content.size) continue
    decorations.push(
      Decoration.widget(header.position, () => repeatedHeader(header), {
        side: -1,
        ignoreSelection: true,
        key: `page-header:${header.position}:${header.height.toFixed(2)}:${header.html.length}`,
      }),
    )
  }

  // No corte entre linhas, o espaçador tem a largura da linha e vem antes do
  // primeiro caractere da que abre a folha: a de cima termina onde já terminava,
  // com a mesma justificação. `vertical-align: top` não soma a descendente ao vão.
  for (const [position, gap] of lines) {
    if (gap <= 0 || position <= 0 || position > view.state.doc.content.size) continue
    decorations.push(
      Decoration.widget(position, () => lineGap(gap), {
        side: -1,
        ignoreSelection: true,
        key: `page-line-gap:${gap.toFixed(2)}`,
      }),
    )
  }

  view.state.doc.descendants((node, offset) => {
    const dx = columns.get(offset)
    if (dx !== undefined && dx !== 0) {
      decorations.push(
        Decoration.node(offset, offset + node.nodeSize, { style: `transform:translateX(${dx}px)` }),
      )
    }

    // O desvio das colunas pode ser negativo; só o zero não se escreve.
    const gap = written.get(offset)
    if (gap === undefined || (gap === 0 && !gaps.has(offset))) return

    decorations.push(
      Decoration.node(offset, offset + node.nodeSize, {
        style: `${node.type.name === 'tableCell' || node.type.name === 'tableHeader' ? 'padding-top' : 'margin-top'}:${gap}px`,
        'data-page-start': 'true',
        'data-page-shift': String(gaps.get(offset) ?? 0),
      }),
    )
  })

  view.dispatch(
    view.state.tr
      .setMeta(paginationKey, DecorationSet.create(view.state.doc, decorations))
      .setMeta('addToHistory', false),
  )
}

/** O que a paginação empurra além dos vãos entre blocos. */
export interface PageGapExtras {
  /** Pela posição do primeiro caractere da linha. */
  readonly lines?: ReadonlyMap<number, number>
  readonly headers?: readonly RepeatedHeader[]
  /** Translação, e não margem: mudar de coluna não muda a altura. */
  readonly columns?: ReadonlyMap<number, number>
}

/** As linhas de cabeçalho de uma tabela, repetidas no alto de uma folha. */
export interface RepeatedHeader {
  /** Início do conteúdo da primeira célula da linha que abre a folha. */
  readonly position: number
  /** Com as colunas da original e só as linhas de cabeçalho. */
  readonly html: string
  readonly height: number
  /** A partir do canto do conteúdo da célula. */
  readonly offsetTop: number
  readonly offsetLeft: number
}

export const REPEATED_HEADER_CLASS = 'page-repeated-header'

function repeatedHeader(header: RepeatedHeader): HTMLElement {
  const element = document.createElement('div')
  element.className = REPEATED_HEADER_CLASS
  element.contentEditable = 'false'
  element.setAttribute('aria-hidden', 'true')
  element.style.cssText =
    `position:absolute;margin-top:${-header.offsetTop}px;margin-left:${-header.offsetLeft}px;` +
    `height:${header.height}px;pointer-events:none;user-select:none;`
  element.innerHTML = header.html
  return element
}

/** `data-page-shift` é o que a medida desconta. */
function lineGap(gap: number): HTMLElement {
  const element = document.createElement('span')
  element.className = LINE_GAP_CLASS
  element.contentEditable = 'false'
  element.setAttribute('aria-hidden', 'true')
  element.dataset.pageShift = String(gap)
  element.style.cssText = `display:inline-block;width:100%;height:${gap}px;vertical-align:top;line-height:0;`
  return element
}

/** Captura ancorada **e** texto; ver `content-styles.ts`. */
export const ANCHOR_TEXT_ATTR = 'data-anchor-text'

export function hasAnchoredImageAndText(node: {
  readonly isTextblock: boolean
  readonly textContent: string
  forEach: (callback: (child: { type: { name: string }; attrs: Record<string, unknown> }) => void) => void
}): boolean {
  if (!node.isTextblock || node.textContent.trim() === '') return false
  let anchored = false
  node.forEach((child) => {
    if (child.type.name === 'image' && child.attrs['anchored'] === true) anchored = true
  })
  return anchored
}

export function isPaginationOnly(transaction: {
  getMeta: (key: PluginKey<DecorationSet>) => unknown
}): boolean {
  return transaction.getMeta(paginationKey) !== undefined
}
