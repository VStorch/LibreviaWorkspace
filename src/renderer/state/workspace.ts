import { create } from 'zustand'
import { DEFAULT_PAGE_SETUP, createEmptyDocument } from '@services/document/model.js'
import { createWorkspaceContext } from './context.js'
import { createDraftActions } from './draft-actions.js'
import { createFileActions } from './file-actions.js'
import { createPrintActions } from './print-actions.js'
import { createSheetActions } from './sheet-actions.js'
import type { WorkspaceState } from './types.js'

export type { DocumentSource, WorkspaceState } from './types.js'

/** Os campos e os ajustes de uma linha; o que tem regra mora nos grupos de ações, ligados por `context.ts`. */
export const useWorkspace = create<WorkspaceState>((set, get) => {
  const ctx = createWorkspaceContext(set, get)
  const empty = createEmptyDocument()

  return {
    file: null,
    page: DEFAULT_PAGE_SETUP,
    initialDoc: empty.doc,
    styles: empty.styles,
    flattened: false,
    beforeReferences: false,
    sections: [],
    beforeSections: false,
    outsideBookmarks: [],
    comments: [],
    commentsOutside: [],
    commentDraft: null,
    beforeComments: false,
    trackChanges: undefined,
    beforeRevisions: false,
    notes: undefined,
    beforeNotes: false,
    beforeMath: false,
    properties: undefined,
    workbook: null,
    generation: 0,
    isDirty: false,
    stats: { characters: 0, words: 0 },
    pageCount: 1,
    recents: [],
    error: null,
    notice: null,
    savedLoss: null,
    pendingDraft: null,
    readOnly: false,
    templateGallery: false,
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
    setPageCount: (pages) => {
      if (get().pageCount !== pages) set({ pageCount: pages })
    },
    setPage: (page) => set({ page, isDirty: true }),
    setSections: (sections) => set({ sections, isDirty: true }),
    setComments: (comments) => set({ comments, isDirty: true }),
    setCommentDraft: (commentDraft) => {
      if (get().commentDraft !== commentDraft) set({ commentDraft })
    },
    setStyles: (styles) => set({ styles, isDirty: true }),
    toggleTrackChanges: () => set({ trackChanges: get().trackChanges !== true, isDirty: true }),
    setProperties: (properties) => set({ properties, isDirty: true }),
    dismissError: () => set({ error: null }),
    dismissNotice: () => set({ notice: null, savedLoss: null }),
    showError: (error) => set({ error }),
    allowEditing: () => set({ readOnly: false }),
    setTemplateGallery: (templateGallery) => set({ templateGallery }),

    ...createFileActions(set, get, ctx),
    ...createSheetActions(set, get),
    ...createDraftActions(set, get, ctx),
    ...createPrintActions(set, get, ctx),
  }
})
