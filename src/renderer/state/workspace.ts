import { create } from 'zustand'
import { DEFAULT_PAGE_SETUP, createEmptyDocument } from '@services/document/model.js'
import { createWorkspaceContext } from './context.js'
import { createDraftActions } from './draft-actions.js'
import { createFileActions } from './file-actions.js'
import { createPrintActions } from './print-actions.js'
import { createSheetActions } from './sheet-actions.js'
import type { WorkspaceState } from './types.js'

export type { DocumentSource, WorkspaceState } from './types.js'

/**
 * O estado do que está aberto na janela.
 *
 * Aqui ficam os campos e os ajustes de uma linha; o que tem regra própria mora
 * num grupo de ações ao lado — arquivo, planilha, rascunho e impressão. Todos
 * conversam pelo mesmo `context.js`, que é onde estão os cinco gestos que os
 * quatro compartilham: chamar o main, ler o que está na tela, esquecer o
 * rascunho, perguntar antes de descartar e pôr um arquivo na tela.
 */
export const useWorkspace = create<WorkspaceState>((set, get) => {
  const ctx = createWorkspaceContext(set, get)

  return {
    file: null,
    page: DEFAULT_PAGE_SETUP,
    initialDoc: createEmptyDocument().doc,
    workbook: null,
    generation: 0,
    isDirty: false,
    stats: { characters: 0, words: 0 },
    estimatedPages: 1,
    recents: [],
    error: null,
    notice: null,
    pendingDraft: null,
    readOnly: false,
    autosaveBroken: false,
    busy: false,

    registerDocumentSource: (source) => ctx.setSource(source),
    markDirty: () => {
      if (!get().isDirty) set({ isDirty: true })
    },
    setStats: (stats) => {
      const current = get().stats
      if (current.characters !== stats.characters || current.words !== stats.words) set({ stats })
    },
    setEstimatedPages: (pages) => {
      if (get().estimatedPages !== pages) set({ estimatedPages: pages })
    },
    setPage: (page) => set({ page, isDirty: true }),
    dismissError: () => set({ error: null }),
    dismissNotice: () => set({ notice: null }),
    showError: (error) => set({ error }),
    allowEditing: () => set({ readOnly: false }),

    ...createFileActions(set, get, ctx),
    ...createSheetActions(set, get),
    ...createDraftActions(set, get, ctx),
    ...createPrintActions(set, get, ctx),
  }
})
