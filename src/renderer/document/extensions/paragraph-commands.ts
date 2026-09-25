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

/**
 * Os estilos do documento aberto, que o editor põe aqui ao recebê-lo.
 *
 * No `storage`, e não nas opções: as extensões são montadas uma vez, e os estilos
 * mudam a cada documento. É deles que sai o valor **que se vê** de um bloco que
 * só carrega a formatação direta.
 */
export interface ParagraphCommandsStorage {
  styles: StyleSheet | null
}

/** O que o bloco vale, com os estilos que o editor conhece. */
export function blockAttrsOf(editor: Editor, node: ProseMirrorNode): Record<string, unknown> {
  return effectiveAttrs(node, stylesOf(editor))
}

/** Sem a extensão montada — num teste, num editor de faixa — não há estilos. */
function stylesOf(editor: Editor): StyleSheet | null {
  return (editor.storage.paragraphCommands as ParagraphCommandsStorage | undefined)?.styles ?? null
}

declare module '@tiptap/core' {
  interface Storage {
    paragraphCommands: ParagraphCommandsStorage
  }

  interface Commands<ReturnType> {
    paragraphCommands: {
      /** Aplica o formulário inteiro aos blocos da seleção. */
      setParagraphFormat: (draft: ParagraphDraft) => ReturnType
      /**
       * Só a entrelinha, para os atalhos `Ctrl+1`, `Ctrl+2` e `Ctrl+5` do Word.
       *
       * Recebe a escolha como o Word a diz — `''` para simples, o fator em linhas
       * (`1.5`) ou a medida (`14pt`) —, e **não** como o CSS a escreve: a
       * conversão depende da fonte de cada bloco, e é `paragraph-format` quem a
       * faz.
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

    /**
     * Os atributos saem de uma função do bloco, e não de um objeto pronto.
     *
     * A entrelinha em CSS depende da altura natural da fonte **daquele** bloco:
     * "1,5 linha" é 1,8311 em Calibri e 1,7249 em Times. Com um valor só para a
     * seleção inteira, um parágrafo de cada fonte receberia a medida do outro.
     */
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

        // Sempre verdadeiro, mesmo sem nada a mudar: "já estava assim" é sucesso,
        // não recusa. Devolver `false` cortava a cadeia do diálogo — e com ela o
        // `focus()`, então clicar em "Aplicar" sem mexer em nada deixava o cursor
        // fora do texto.
        return true
      }

    return {
      setParagraphFormat: (draft) =>
        applyAttrs((node) => ({ ...paragraphAttrsFrom(draft, node.attrs, effective(node)) })),
      // Contra o valor efetivo: a conversão precisa da fonte que o bloco **usa**,
      // e ela pode vir do estilo.
      setBlockLineHeight: (value) =>
        applyAttrs((node) => ({ lineHeight: lineHeightAttrFrom(value, effective(node)) })),
    }
  },
})

/**
 * A entrelinha do bloco sob o cursor, como o seletor rápido da barra a mostra.
 *
 * Em linhas do Word, e não na medida do CSS: o atributo de um parágrafo de
 * Calibri a 1,5 linha é 1,8311, e a barra mostrava esse número. O simples volta
 * como vazio porque na barra a opção se chama "Simples" e é a primeira da lista.
 */
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
