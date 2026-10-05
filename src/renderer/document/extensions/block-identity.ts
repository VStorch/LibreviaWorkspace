import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state'

/**
 * The `oid` the reader stamps on each block (`BodyReader.NewBlock`) decides what saving does
 * **not** rewrite (`DocxWriter.OidOf`). ProseMirror drops attributes outside the schema: without
 * this extension saving would silently become a full regeneration. It also goes into the HTML
 * (`data-oid`), to survive cut, paste and undo.
 */

export interface BlockIdentityOptions {
  types: string[]
}

/**
 * Two blocks with the same `oid` (Enter and paste duplicate it) would make the writer regenerate
 * the second. The first keeps the identity, and the new one starts without it. Only when there is
 * something to fix, and without descending into paragraphs.
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
    // The nodes where `BodyReader` calls `NewBlock`; a list item is a `w:p`, and the table of
    // contents a `w:sdt`.
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
           * An image or box outside the flow, opaque data like `oid`. On the same nodes: a
           * paragraph with an image and a page break becomes `pageBreak`. Not in the HTML: whoever
           * draws reads it from the model.
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
