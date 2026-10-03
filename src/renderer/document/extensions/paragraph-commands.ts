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

/** O formulário inteiro numa transação: oito em fila pediriam oito `Ctrl+Z`. */

/** No OOXML isto é propriedade do bloco. */
const DEFAULT_TYPES: readonly string[] = ['paragraph', 'heading', 'bulletList', 'orderedList']

export interface ParagraphCommandsOptions {
  types: string[]
}

/** No `storage`, e não nas opções: as extensões são montadas uma vez, e os estilos mudam a cada documento. */
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
      /** Para `Ctrl+1`, `Ctrl+2` e `Ctrl+5`, como o Word diz: a conversão em CSS depende da fonte de cada bloco. */
      setBlockLineHeight: (value: string) => ReturnType
    }
  }
}

/** O primeiro bloco, como no Word: uma média inventaria um número que não é de nenhum. */
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

    /** Por bloco: "1,5 linha" é 1,8311 em Calibri e 1,7249 em Times. */
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
            // A lista não tem `textAlign`: o alinhamento é dos itens.
            if (!(name in declared)) continue
            if (node.attrs[name] === value) continue

            tr.setNodeAttribute(pos, name, value)
            changed = true
          }

          return true
        })

        if (changed && dispatch !== undefined) dispatch(tr)

        // "Já estava assim" é sucesso: `false` cortaria a cadeia, e o `focus()` com ela.
        return true
      }

    return {
      setParagraphFormat: (draft) =>
        applyAttrs((node) => ({ ...paragraphAttrsFrom(draft, node.attrs, effective(node)) })),
      // Contra o valor efetivo: a fonte pode vir do estilo.
      setBlockLineHeight: (value) =>
        applyAttrs((node) => ({ lineHeight: lineHeightAttrFrom(value, effective(node)) })),
    }
  },
})

/** Em linhas do Word; o simples volta vazio, a primeira opção da barra. */
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
