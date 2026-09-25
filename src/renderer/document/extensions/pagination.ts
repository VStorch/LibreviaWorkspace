import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import { LINE_GAP_CLASS } from '../line-boxes.js'

/**
 * O vão entre uma folha e a seguinte.
 *
 * Paginar num editor de texto é empurrar: o bloco que abre uma página nova
 * ganha uma margem superior do tamanho exato do que sobrou da folha anterior
 * mais as duas margens e o vão entre os papéis. O texto continua sendo um
 * fluxo só, e as folhas brancas são desenhadas atrás, nas posições que essa
 * conta produz.
 *
 * A margem entra como **decoração de nó**, e não como um elemento inserido no
 * meio do texto. É a diferença que decide o resto: um espaçador de verdade
 * dentro do `contenteditable` entraria na seleção, no `Ctrl+A` e no que a
 * pessoa copia — colar um trecho de duas páginas levaria junto um pedaço de
 * papel. Decoração não existe para o documento: é aparência aplicada sobre
 * nós que continuam intactos, e some sem deixar rastro quando a página muda de
 * lugar.
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

            // Enquanto a medição não chega, as decorações acompanham a edição:
            // sem isto, digitar no meio do documento deslocaria as folhas de
            // baixo até a próxima medida, e elas piscariam de lugar.
            return current.map(transaction.mapping, transaction.doc)
          },
        },
        props: {
          decorations: (state) => paginationKey.getState(state),
        },
      }),
    ]
  },
})

/**
 * Aplica os vãos por posição de nó, inclusive itens e células dentro de blocos.
 *
 * A transação não entra no histórico: desfazer precisa voltar o que a pessoa
 * escreveu, não o lugar onde a página caiu.
 */
export function applyPageGaps(
  view: EditorView,
  written: ReadonlyMap<number, number>,
  gaps: ReadonlyMap<number, number>,
  /** Vãos entre linhas de um parágrafo cortado, pela posição do primeiro caractere da linha. */
  lines: ReadonlyMap<number, number> = new Map(),
): void {
  const decorations: Decoration[] = []

  // O corte no meio do parágrafo não tem nó a empurrar: o espaçador é um
  // elemento da largura da linha, antes do primeiro caractere da linha que abre
  // a folha. Ele cabe só numa linha própria, então a linha de cima termina onde
  // já terminava — com a mesma justificação, porque a quebra continua sendo
  // automática e não forçada — e a de baixo recomeça no topo da folha seguinte.
  // `vertical-align: top` faz a linha dele ter a altura dele, sem somar a
  // descendente do texto ao vão.
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
    const gap = written.get(offset)
    if (gap === undefined || gap <= 0) return

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

/** O espaçador de um corte entre linhas; `data-page-shift` é o que a medida desconta. */
function lineGap(gap: number): HTMLElement {
  const element = document.createElement('span')
  element.className = LINE_GAP_CLASS
  element.contentEditable = 'false'
  element.setAttribute('aria-hidden', 'true')
  element.dataset.pageShift = String(gap)
  element.style.cssText = `display:inline-block;width:100%;height:${gap}px;vertical-align:top;line-height:0;`
  return element
}

/** A transação só mexeu em paginação — não é edição do documento. */
export function isPaginationOnly(transaction: {
  getMeta: (key: PluginKey<DecorationSet>) => unknown
}): boolean {
  return transaction.getMeta(paginationKey) !== undefined
}
