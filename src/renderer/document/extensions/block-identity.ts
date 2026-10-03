import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state'

/**
 * O `oid` que o leitor carimba em cada bloco (`BodyReader.NewBlock`) decide o que
 * a gravação **não** reescreve (`DocxWriter.OidOf`). O ProseMirror descarta
 * atributo fora do schema: sem esta extensão a gravação viraria regeneração
 * completa, em silêncio. Vai também ao HTML (`data-oid`), para atravessar
 * recortar, colar e desfazer.
 */

export interface BlockIdentityOptions {
  types: string[]
}

/**
 * Dois blocos com o mesmo `oid` (o Enter e a colagem duplicam) fariam o gravador
 * regenerar o segundo. O primeiro fica com a identidade, e o novo nasce sem.
 * Só quando há o que corrigir, e sem descer dentro do parágrafo.
 */
export function uniqueOids(): Plugin {
  return new Plugin({
    key: new PluginKey('blockIdentityUnique'),

    appendTransaction(transactions, _oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) return null

      const seen = new Set<string>()
      let corrections: Transaction | null = null

      newState.doc.descendants((node, position) => {
        const oid: unknown = node.attrs['oid']
        if (typeof oid === 'string' && oid.length > 0) {
          if (seen.has(oid)) {
            corrections ??= newState.tr
            corrections.setNodeAttribute(position, 'oid', null)
          } else {
            seen.add(oid)
          }
        }

        return !node.isTextblock
      })

      return corrections
    },
  })
}

export const BlockIdentity = Extension.create<BlockIdentityOptions>({
  name: 'blockIdentity',

  addProseMirrorPlugins() {
    return [uniqueOids()]
  },

  addOptions() {
    // Os nós em que `BodyReader` chama `NewBlock`; o item de lista é um `w:p`, e o sumário, o `w:sdt`.
    return { types: ['paragraph', 'heading', 'pageBreak', 'listItem', 'table', 'tableOfContents'] }
  },

  addGlobalAttributes() {
    return [
      {
        types: this.options.types,
        attributes: {
          oid: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-oid'),
            renderHTML: (attributes) => {
              const oid = attributes['oid']
              return typeof oid === 'string' && oid.length > 0 ? { 'data-oid': oid } : {}
            },
          },

          /**
           * Imagem ou caixa fora do fluxo, dado opaco como o `oid`. Nos mesmos nós:
           * o parágrafo com imagem e quebra de página vira `pageBreak`. Não vai ao
           * HTML: quem desenha lê do modelo.
           */
          floats: {
            default: null,
            parseHTML: () => null,
            renderHTML: () => ({}),
          },
        },
      },
    ]
  },
})
