import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, TextSelection, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { findOccurrences, stepIndex } from '@services/document/search.js'

/**
 * Localizar e substituir.
 *
 * Não há extensão oficial para isto no Tiptap, então é nossa. A lógica de
 * casamento vive em `@services/document/search.ts` (pura e testada); aqui fica
 * só o que precisa do ProseMirror: mapear posições, desenhar os destaques e
 * aplicar as substituições.
 *
 * A busca é feita por bloco de texto usando `textBetween`, e não nó a nó,
 * porque uma palavra pode estar partida em vários nós de texto quando parte
 * dela tem formatação — "Ne**gr**ito" são três nós e uma palavra só.
 */

export interface SearchMatch {
  readonly from: number
  readonly to: number
}

export interface SearchStatus {
  readonly total: number
  /** Posição da ocorrência atual, começando em 1. Zero quando não há nenhuma. */
  readonly current: number
}

interface SearchPluginState {
  term: string
  caseSensitive: boolean
  matches: SearchMatch[]
  currentIndex: number
  decorations: DecorationSet
}

export interface SearchReplaceOptions {
  onStatusChange: (status: SearchStatus) => void
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    searchReplace: {
      setSearchTerm: (term: string, caseSensitive: boolean) => ReturnType
      clearSearch: () => ReturnType
      goToMatch: (delta: number) => ReturnType
      replaceCurrentMatch: (replacement: string) => ReturnType
      replaceAllMatches: (replacement: string) => ReturnType
    }
  }
}

export const searchPluginKey = new PluginKey<SearchPluginState>('searchReplace')

function collectMatches(doc: ProseMirrorNode, term: string, caseSensitive: boolean): SearchMatch[] {
  if (term.length === 0) return []

  const matches: SearchMatch[] = []

  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true

    // O separador de um caractere para nós folha mantém o comprimento do texto
    // alinhado com as posições do documento.
    const text = node.textBetween(0, node.content.size, undefined, ' ')

    for (const occurrence of findOccurrences(text, term, caseSensitive)) {
      matches.push({ from: pos + 1 + occurrence.start, to: pos + 1 + occurrence.end })
    }

    // Um bloco de texto não contém outro; descer seria reprocessar o conteúdo.
    return false
  })

  return matches
}

function buildDecorations(doc: ProseMirrorNode, state: SearchPluginState): DecorationSet {
  if (state.matches.length === 0) return DecorationSet.empty

  const decorations = state.matches.map((match, index) =>
    Decoration.inline(match.from, match.to, {
      class: index === state.currentIndex ? 'search-hit search-hit--current' : 'search-hit',
    }),
  )

  return DecorationSet.create(doc, decorations)
}

function recompute(state: SearchPluginState, doc: ProseMirrorNode): SearchPluginState {
  const matches = collectMatches(doc, state.term, state.caseSensitive)

  let currentIndex = state.currentIndex
  if (matches.length === 0) currentIndex = -1
  else if (currentIndex < 0 || currentIndex >= matches.length) currentIndex = 0

  const next: SearchPluginState = { ...state, matches, currentIndex, decorations: DecorationSet.empty }
  return { ...next, decorations: buildDecorations(doc, next) }
}

export const SearchReplace = Extension.create<SearchReplaceOptions>({
  name: 'searchReplace',

  addOptions() {
    return { onStatusChange: () => undefined }
  },

  addProseMirrorPlugins() {
    const notify = (status: SearchStatus): void => this.options.onStatusChange(status)

    return [
      new Plugin<SearchPluginState>({
        key: searchPluginKey,

        state: {
          init: () => ({
            term: '',
            caseSensitive: false,
            matches: [],
            currentIndex: -1,
            decorations: DecorationSet.empty,
          }),

          apply(tr, value, _oldState, newState) {
            const meta = tr.getMeta(searchPluginKey) as Partial<SearchPluginState> | undefined

            if (meta === undefined && !tr.docChanged) return value

            const merged: SearchPluginState = { ...value, ...meta }
            return recompute(merged, newState.doc)
          },
        },

        view() {
          let lastReported = ''
          return {
            update(view) {
              const state = searchPluginKey.getState(view.state)
              if (state === undefined) return

              const status: SearchStatus = {
                total: state.matches.length,
                current: state.currentIndex < 0 ? 0 : state.currentIndex + 1,
              }
              const signature = `${status.total}/${status.current}`
              if (signature === lastReported) return

              lastReported = signature
              notify(status)
            },
          }
        },

        props: {
          decorations: (state) => searchPluginKey.getState(state)?.decorations ?? DecorationSet.empty,
        },
      }),
    ]
  },

  addCommands() {
    /** Leva o cursor até a ocorrência para que ela role para a área visível. */
    const revealMatch = (tr: Transaction, match: SearchMatch): void => {
      tr.setSelection(TextSelection.create(tr.doc, match.from, match.to))
      tr.scrollIntoView()
    }

    return {
      setSearchTerm:
        (term, caseSensitive) =>
        ({ tr, dispatch }) => {
          if (dispatch !== undefined) {
            tr.setMeta(searchPluginKey, { term, caseSensitive, currentIndex: 0 })
            dispatch(tr)
          }
          return true
        },

      clearSearch:
        () =>
        ({ tr, dispatch }) => {
          if (dispatch !== undefined) {
            tr.setMeta(searchPluginKey, { term: '', currentIndex: -1 })
            dispatch(tr)
          }
          return true
        },

      goToMatch:
        (delta) =>
        ({ tr, state, dispatch }) => {
          const pluginState = searchPluginKey.getState(state)
          if (pluginState === undefined || pluginState.matches.length === 0) return false

          const next = stepIndex(pluginState.currentIndex, pluginState.matches.length, delta)
          if (dispatch !== undefined) {
            tr.setMeta(searchPluginKey, { currentIndex: next })
            const match = pluginState.matches[next]
            if (match !== undefined) revealMatch(tr, match)
            dispatch(tr)
          }
          return true
        },

      replaceCurrentMatch:
        (replacement) =>
        ({ tr, state, dispatch }) => {
          const pluginState = searchPluginKey.getState(state)
          const match = pluginState?.matches[pluginState.currentIndex]
          if (match === undefined) return false

          if (dispatch !== undefined) {
            tr.insertText(replacement, match.from, match.to)
            // O índice fica onde está: depois de substituir, a próxima
            // ocorrência assume a mesma posição na lista.
            tr.setMeta(searchPluginKey, {})
            dispatch(tr)
          }
          return true
        },

      replaceAllMatches:
        (replacement) =>
        ({ tr, state, dispatch }) => {
          const pluginState = searchPluginKey.getState(state)
          if (pluginState === undefined || pluginState.matches.length === 0) return false

          if (dispatch !== undefined) {
            // De trás para frente: substituir da esquerda para a direita
            // invalidaria as posições ainda não processadas.
            for (const match of [...pluginState.matches].reverse()) {
              tr.insertText(replacement, match.from, match.to)
            }
            tr.setMeta(searchPluginKey, { currentIndex: -1 })
            dispatch(tr)
          }
          return true
        },
    }
  },
})
