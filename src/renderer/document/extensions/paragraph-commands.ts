import { Extension, type CommandProps, type Editor } from '@tiptap/core'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  paragraphAttrsFrom,
  paragraphDraftFrom,
  type ParagraphDraft,
} from '@services/document/paragraph-format.js'

/**
 * Ler e escrever a formatação de parágrafo dos blocos selecionados.
 *
 * Os atributos já existem: espaçamento, entrelinha, recuo e "manter com o
 * próximo" são de `BlockFormat`, e o alinhamento é de `TextAlign`. O que não
 * existia era um comando que os escrevesse **todos de uma vez** nos blocos que a
 * seleção cobre. Sem ele o diálogo de parágrafo despacharia oito transações em
 * fila, e desfazer pediria oito `Ctrl+Z`.
 *
 * Extensão separada de `BlockFormat` porque são responsabilidades diferentes: lá
 * mora o que o atributo **é** e como ele vira CSS; aqui, o que a interface faz
 * com ele.
 */

/** Os mesmos tipos de `BlockFormat`: no OOXML isto é propriedade do bloco. */
const DEFAULT_TYPES: readonly string[] = ['paragraph', 'heading', 'bulletList', 'orderedList']

export interface ParagraphCommandsOptions {
  types: string[]
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    paragraphCommands: {
      /** Aplica o formulário inteiro aos blocos da seleção. */
      setParagraphFormat: (draft: ParagraphDraft) => ReturnType
      /**
       * Só a entrelinha, para os atalhos `Ctrl+1`, `Ctrl+2` e `Ctrl+5` do Word.
       *
       * Recebe a medida como o CSS a escreve, que é a forma em que o atributo
       * vive: `normal`, um fator, ou pontos.
       */
      setBlockLineHeight: (value: string) => ReturnType
    }
  }
}

/**
 * O primeiro bloco que a seleção toca, como o diálogo o mostra.
 *
 * O primeiro, e não uma média dos vários: com dois parágrafos de espaçamento
 * diferente selecionados, qualquer fusão inventaria um número que não é de
 * nenhum deles. O Word faz assim — mostra o do primeiro e aplica a todos.
 */
export function paragraphDraftAt(editor: Editor, types: readonly string[] = DEFAULT_TYPES): ParagraphDraft {
  const { from, to } = editor.state.selection
  let attrs: Record<string, unknown> | null = null

  editor.state.doc.nodesBetween(from, to, (node) => {
    if (attrs !== null) return false
    if (types.includes(node.type.name)) attrs = node.attrs
    return true
  })

  return attrs === null ? DEFAULT_PARAGRAPH_DRAFT : paragraphDraftFrom(attrs)
}

export const ParagraphCommands = Extension.create<ParagraphCommandsOptions>({
  name: 'paragraphCommands',

  addOptions() {
    return { types: [...DEFAULT_TYPES] }
  },

  addCommands() {
    const types = this.options.types

    const applyAttrs =
      (attrs: Record<string, unknown>) =>
      ({ state, tr, dispatch }: CommandProps): boolean => {
        const { from, to } = state.selection
        let changed = false

        state.doc.nodesBetween(from, to, (node, pos) => {
          if (!types.includes(node.type.name)) return true

          const declared = node.type.spec.attrs
          if (declared === undefined) return true

          for (const [name, value] of Object.entries(attrs)) {
            // Atributo que o tipo do bloco não declara faria o ProseMirror
            // reclamar: a lista não tem `textAlign`, porque o alinhamento dela é
            // dos itens.
            if (!(name in declared)) continue
            if (node.attrs[name] === value) continue

            tr.setNodeAttribute(pos, name, value)
            changed = true
          }

          return true
        })

        if (changed && dispatch !== undefined) dispatch(tr)
        return changed
      }

    return {
      setParagraphFormat: (draft) => applyAttrs({ ...paragraphAttrsFrom(draft) }),
      setBlockLineHeight: (value) => applyAttrs({ lineHeight: value }),
    }
  },
})

/**
 * A entrelinha do bloco sob o cursor, como o seletor rápido da barra a mostra.
 *
 * `normal` volta como vazio porque na barra a opção se chama "Simples" e é a
 * primeira da lista — é a mesma coisa dita em duas línguas, a do CSS e a do Word.
 */
export function blockLineHeightOf(editor: Editor, types: readonly string[] = DEFAULT_TYPES): string {
  const { from, to } = editor.state.selection
  let value: string | null = null

  editor.state.doc.nodesBetween(from, to, (node) => {
    if (value !== null) return false
    if (!types.includes(node.type.name)) return true

    const declared = node.attrs['lineHeight']
    value = typeof declared === 'string' ? declared : ''
    return true
  })

  return value === null || value === 'normal' ? '' : value
}
