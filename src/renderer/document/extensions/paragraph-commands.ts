import { Extension, type CommandProps, type Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  lineHeightAttrFrom,
  lineSpacingChoiceOf,
  paragraphAttrsFrom,
  paragraphDraftFrom,
  type ParagraphDraft,
} from '@services/document/paragraph-format.js'
import { effectiveAttrs } from '@services/document/style-cascade.js'
import type { StyleSheet } from '@services/document/styles.js'

/** The whole form in one transaction: eight in a row would take eight `Ctrl+Z`. */

/** In OOXML this is a block property. */
const DEFAULT_TYPES: readonly string[] = ['paragraph', 'heading', 'bulletList', 'orderedList']

export interface ParagraphCommandsOptions {
  types: string[]
}

/** In `storage`, not options: extensions are mounted once, and styles change with each document. */
export interface ParagraphCommandsStorage {
  styles: StyleSheet | null
}

export function blockAttrsOf(editor: Editor, node: ProseMirrorNode): Record<string, unknown> {
  return effectiveAttrs(node, stylesOf(editor))
}

function stylesOf(editor: Editor): StyleSheet | null {
  return (editor.storage.paragraphCommands as ParagraphCommandsStorage | undefined)?.styles ?? null
}

declare module '@tiptap/core' {
  interface Storage {
    paragraphCommands: ParagraphCommandsStorage
  }

  interface Commands<ReturnType> {
    paragraphCommands: {
      setParagraphFormat: (draft: ParagraphDraft) => ReturnType
      /**
       * For `Ctrl+1`, `Ctrl+2` and `Ctrl+5`, as Word says: the CSS conversion depends on each
       * block's font.
       */
      setBlockLineHeight: (value: string) => ReturnType
    }
  }
}

/** The first block, as in Word: an average would invent a number that belongs to none. */
export function paragraphDraftAt(editor: Editor, types: readonly string[] = DEFAULT_TYPES): ParagraphDraft {
  const { from, to } = editor.state.selection
  let attrs: Record<string, unknown> | null = null

  editor.state.doc.nodesBetween(from, to, (node) => {
    if (attrs !== null) return false
    if (types.includes(node.type.name)) attrs = blockAttrsOf(editor, node)
    return true
  })

  return attrs === null ? DEFAULT_PARAGRAPH_DRAFT : paragraphDraftFrom(attrs)
}

export const ParagraphCommands = Extension.create<ParagraphCommandsOptions, ParagraphCommandsStorage>({
  name: 'paragraphCommands',

  addOptions() {
    return { types: [...DEFAULT_TYPES] }
  },

  addStorage() {
    return { styles: null }
  },

  addCommands() {
    const types = this.options.types
    const effective = (node: ProseMirrorNode): Record<string, unknown> =>
      effectiveAttrs(node, this.storage.styles)

    /** Per block: "1.5 lines" is 1.8311 in Calibri and 1.7249 in Times. */
    const applyAttrs =
      (attrsOf: (node: ProseMirrorNode) => Record<string, unknown>) =>
      ({ state, tr, dispatch }: CommandProps): boolean => {
        const { from, to } = state.selection
        let changed = false

        state.doc.nodesBetween(from, to, (node, pos) => {
          if (!types.includes(node.type.name)) return true

          const declared = node.type.spec.attrs
          if (declared === undefined) return true

          for (const [name, value] of Object.entries(attrsOf(node))) {
            // A list has no `textAlign`: alignment belongs to the items.
            if (!(name in declared)) continue
            if (node.attrs[name] === value) continue

            tr.setNodeAttribute(pos, name, value)
            changed = true
          }

          return true
        })

        if (changed && dispatch !== undefined) dispatch(tr)

        // "Already like that" is success: `false` would cut the chain, and `focus()` with it.
        return true
      }

    return {
      setParagraphFormat: (draft) =>
        applyAttrs((node) => ({ ...paragraphAttrsFrom(draft, node.attrs, effective(node)) })),
      // Against the effective value: the font may come from the style.
      setBlockLineHeight: (value) =>
        applyAttrs((node) => ({ lineHeight: lineHeightAttrFrom(value, effective(node)) })),
    }
  },
})

/** In Word lines; single comes back empty, the toolbar's first option. */
export function blockLineHeightOf(editor: Editor, types: readonly string[] = DEFAULT_TYPES): string {
  const { from, to } = editor.state.selection
  let value: string | null = null

  editor.state.doc.nodesBetween(from, to, (node) => {
    if (value !== null) return false
    if (!types.includes(node.type.name)) return true

    value = lineSpacingChoiceOf(blockAttrsOf(editor, node))
    return true
  })

  return value ?? ''
}
