import type { StoreApi } from 'zustand'
import { AppError, type SerializedError } from '@shared/errors.js'
import { DiscardChoice } from '@shared/types.js'
import type { IpcResult } from '@shared/ipc.js'
import type { DocumentModel } from '@services/document/model.js'
import { onlyDefined, onlyNonEmpty, onlyTrue, serializeDocument } from '@services/document/serialize.js'
import { commentAnchorIdsOfJson, commentsOutsideOf, resolveComments } from '@services/document/comments.js'
import { marksOfJson, resolveSections } from '@services/document/sections.js'
import { serializeWorkbook } from '@services/spreadsheet/serialize.js'
import { t } from '../i18n.js'
import type { DocumentSource, LoadedFile, OpenFile, WorkspaceState } from './types.js'

export type SetWorkspace = StoreApi<WorkspaceState>['setState']
export type GetWorkspace = StoreApi<WorkspaceState>['getState']

/** Steps shared by open, save, recover and print, written once so they do not drift. */
export interface WorkspaceContext {
  call: <T>(operation: () => Promise<IpcResult<T>>) => Promise<T | null>
  currentModel: () => DocumentModel
  /** In the internal format, the same as the draft. */
  currentContent: () => string
  forgetDraft: () => Promise<void>
  /** `false` when the user gave up: the caller stops without changing anything. */
  ensureChangesHandled: () => Promise<boolean>
  show: (loaded: LoadedFile) => void
  source: () => DocumentSource | null
  setSource: (source: DocumentSource | null) => void
}

export function toSerialized(cause: unknown): SerializedError {
  if (cause instanceof AppError) return cause.toSerialized()
  return { code: 'INTERNAL', message: t('shell.error.unexpected') }
}

export function createWorkspaceContext(set: SetWorkspace, get: GetWorkspace): WorkspaceContext {
  let documentSource: DocumentSource | null = null

  async function call<T>(operation: () => Promise<IpcResult<T>>): Promise<T | null> {
    set({ busy: true })
    try {
      const result = await operation()
      if (result.ok) return result.data
      set({ error: result.error })
      return null
    } catch {
      set({
        error: { code: 'INTERNAL', message: t('shell.error.communicationFailed') },
      })
      return null
    } finally {
      set({ busy: false })
    }
  }

  function currentModel(): DocumentModel {
    const state = get()
    return modelOf(state, documentSource?.readDoc() ?? state.initialDoc)
  }

  function currentContent(): string {
    const { workbook } = get()
    return workbook === null ? serializeDocument(currentModel()) : serializeWorkbook(workbook)
  }

  /**
   * After saving and when switching files. Failure is swallowed: a stale draft only costs a prompt.
   */
  async function forgetDraft(): Promise<void> {
    set({ autosaveBroken: false })
    await window.api.recovery.discard({}).catch(() => undefined)
  }

  async function ensureChangesHandled(): Promise<boolean> {
    const state = get()
    if (!state.isDirty) return true

    const answer = await call(() =>
      window.api.dialog.confirmDiscard({ fileName: state.file?.name ?? t('shell.file.untitled') }),
    )
    if (answer === null) return false

    if (answer.choice === DiscardChoice.Cancel) return false
    if (answer.choice === DiscardChoice.Discard) return true
    return get().save()
  }

  function show({ file, model, workbook }: LoadedFile): void {
    // Also on recovery: the content is already on screen, and autosave writes it again.
    void forgetDraft()

    set((state) => loadedState(file, model, workbook, state.generation))
  }

  return {
    call,
    currentModel,
    currentContent,
    forgetDraft,
    ensureChangesHandled,
    show,
    source: () => documentSource,
    setSource: (source) => {
      documentSource = source
    },
  }
}

function modelOf(state: WorkspaceState, read: WorkspaceState['initialDoc']): DocumentModel {
  // The sections the text uses, in its order (`resolveSections`); the document attribute does not
  // go to the file.
  const resolved = resolveSections(marksOfJson(read), read.attrs?.['bodySection'], state.page, state.sections)
  const { bodySection: _bodySection, ...docAttrs } = read.attrs ?? {}
  void _bodySection
  const doc = read.attrs === undefined ? read : { ...read, attrs: docAttrs }
  // Only the comments the text supports; a draft older than comments goes as it came.
  const comments = state.beforeComments
    ? state.comments
    : resolveComments(commentAnchorIdsOfJson(read), state.comments, new Set(state.commentsOutside))
  return {
    page: resolved.page,
    doc,
    styles: state.styles,
    ...onlyTrue('flattened', state.flattened),
    ...onlyTrue('beforeReferences', state.beforeReferences),
    ...onlyNonEmpty('sections', resolved.sections),
    ...onlyTrue('beforeSections', state.beforeSections),
    ...onlyNonEmpty('outsideBookmarks', state.outsideBookmarks),
    ...onlyNonEmpty('comments', comments),
    ...onlyTrue('beforeComments', state.beforeComments),
    ...onlyDefined('trackChanges', state.trackChanges),
    ...onlyTrue('beforeRevisions', state.beforeRevisions),
    ...onlyDefined('notes', state.notes),
    ...onlyTrue('beforeNotes', state.beforeNotes),
    ...onlyTrue('beforeMath', state.beforeMath),
    ...onlyDefined('properties', state.properties),
  }
}

/** The state of a just-opened file; without a file, the home screen's. */
export function loadedState(
  file: OpenFile | null,
  model: DocumentModel,
  workbook: LoadedFile['workbook'],
  generation: number,
): Partial<WorkspaceState> {
  return {
    file,
    workbook,
    page: model.page,
    initialDoc: model.doc,
    styles: model.styles,
    flattened: model.flattened === true,
    beforeReferences: model.beforeReferences === true,
    sections: model.sections ?? [],
    beforeSections: model.beforeSections === true,
    outsideBookmarks: model.outsideBookmarks ?? [],
    comments: model.comments ?? [],
    commentsOutside: commentsOutsideOf(model.doc, model.comments ?? []),
    commentDraft: null,
    beforeComments: model.beforeComments === true,
    trackChanges: model.trackChanges,
    beforeRevisions: model.beforeRevisions === true,
    notes: model.notes,
    beforeNotes: model.beforeNotes === true,
    beforeMath: model.beforeMath === true,
    properties: model.properties,
    generation: generation + 1,
    isDirty: false,
    error: null,
    // The notice and the lock belong to this file, and die with it.
    notice: null,
    savedLoss: null,
    readOnly: false,
  }
}
