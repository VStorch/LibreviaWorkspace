import { Extension, type CommandProps } from '@tiptap/core'

/**
 * Recuo de parágrafo.
 *
 * Não existe extensão oficial para isto, então é nossa. O recuo é guardado
 * como um número de níveis — e não como uma medida em CSS — porque a Fase 4
 * precisa traduzi-lo para `w:ind` do DOCX, que também trabalha em passos.
 *
 * O atalho é `Ctrl+]` / `Ctrl+[` em vez de `Tab`: dentro de uma lista, `Tab`
 * já significa "aninhar item", e roubar essa tecla quebraria as listas.
 */

export const INDENT_STEP_EM = 2.5
export const MAX_INDENT_LEVEL = 10

export interface IndentOptions {
  types: string[]
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    indent: {
      indent: () => ReturnType
      outdent: () => ReturnType
    }
  }
}

export function clampIndentLevel(level: number): number {
  if (!Number.isFinite(level)) return 0
  return Math.min(Math.max(Math.round(level), 0), MAX_INDENT_LEVEL)
}

export const Indent = Extension.create<IndentOptions>({
  name: 'indent',

  addOptions() {
    return { types: ['paragraph', 'heading'] }
  },

  addGlobalAttributes() {
    return [
      {
        types: this.options.types,
        attributes: {
          indent: {
            default: 0,
            parseHTML: (element) => {
              const margin = element.style.marginLeft
              if (margin === '') return 0
              return clampIndentLevel(Number.parseFloat(margin) / INDENT_STEP_EM)
            },
            renderHTML: (attributes) => {
              const level = clampIndentLevel(Number(attributes['indent'] ?? 0))
              return level === 0 ? {} : { style: `margin-left: ${level * INDENT_STEP_EM}em` }
            },
          },
        },
      },
    ]
  },

  addCommands() {
    const types = this.options.types

    const shiftBy =
      (delta: number) =>
      () =>
      ({ state, tr, dispatch }: CommandProps): boolean => {
        const { from, to } = state.selection
        let changed = false

        state.doc.nodesBetween(from, to, (node, pos) => {
          if (!types.includes(node.type.name)) return true

          const current = clampIndentLevel(Number(node.attrs['indent'] ?? 0))
          const next = clampIndentLevel(current + delta)
          if (next !== current) {
            tr.setNodeAttribute(pos, 'indent', next)
            changed = true
          }
          return true
        })

        if (changed && dispatch !== undefined) dispatch(tr)
        return changed
      }

    return { indent: shiftBy(1), outdent: shiftBy(-1) }
  },

  addKeyboardShortcuts() {
    return {
      'Mod-]': () => this.editor.commands.indent(),
      'Mod-[': () => this.editor.commands.outdent(),
    }
  },
})
