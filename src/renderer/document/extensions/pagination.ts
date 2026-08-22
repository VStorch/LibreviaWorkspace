import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'

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
 * Aplica os vãos medidos, por índice de bloco de primeiro nível.
 *
 * A transação não entra no histórico: desfazer precisa voltar o que a pessoa
 * escreveu, não o lugar onde a página caiu.
 */
export function applyPageGaps(view: EditorView, gaps: ReadonlyMap<number, number>): void {
  const decorations: Decoration[] = []

  view.state.doc.forEach((node, offset, index) => {
    const gap = gaps.get(index)
    if (gap === undefined || gap <= 0) return

    decorations.push(
      Decoration.node(offset, offset + node.nodeSize, {
        style: `margin-top:${Math.round(gap)}px`,
        'data-page-start': 'true',
      }),
    )
  })

  view.dispatch(
    view.state.tr
      .setMeta(paginationKey, DecorationSet.create(view.state.doc, decorations))
      .setMeta('addToHistory', false),
  )
}

/** A transação só mexeu em paginação — não é edição do documento. */
export function isPaginationOnly(transaction: {
  getMeta: (key: PluginKey<DecorationSet>) => unknown
}): boolean {
  return transaction.getMeta(paginationKey) !== undefined
}
