import type { StoreApi } from 'zustand'
import { AppError, type SerializedError } from '@shared/errors.js'
import { DiscardChoice } from '@shared/types.js'
import type { IpcResult } from '@shared/ipc.js'
import type { DocumentModel } from '@services/document/model.js'
import { serializeDocument } from '@services/document/serialize.js'
import { commentAnchorIdsOfJson, commentsOutsideOf, resolveComments } from '@services/document/comments.js'
import { marksOfJson, resolveSections } from '@services/document/sections.js'
import { serializeWorkbook } from '@services/spreadsheet/serialize.js'
import { t } from '../i18n.js'
import type { DocumentSource, LoadedFile, OpenFile, WorkspaceState } from './types.js'

export type SetWorkspace = StoreApi<WorkspaceState>['setState']
export type GetWorkspace = StoreApi<WorkspaceState>['getState']

/** Gestos comuns a abrir, salvar, recuperar e imprimir, escritos uma vez para não divergirem. */
export interface WorkspaceContext {
  call: <T>(operation: () => Promise<IpcResult<T>>) => Promise<T | null>
  currentModel: () => DocumentModel
  /** No formato interno, o mesmo do rascunho. */
  currentContent: () => string
  forgetDraft: () => Promise<void>
  /** `false` quando o usuário desistiu: quem chamou para sem alterar nada. */
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

  /** Depois de gravar e ao trocar de arquivo. A falha é engolida: um rascunho velho custa só um aviso. */
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
    // Vale também na recuperação: o conteúdo já está na tela, e o autosave o grava de novo.
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
  // As seções que o texto usa, na ordem dele (`resolveSections`); o atributo do documento não vai ao arquivo.
  const resolved = resolveSections(marksOfJson(read), read.attrs?.['bodySection'], state.page, state.sections)
  const { bodySection: _bodySection, ...docAttrs } = read.attrs ?? {}
  void _bodySection
  const doc = read.attrs === undefined ? read : { ...read, attrs: docAttrs }
  // Só os comentários que o texto sustenta; o rascunho anterior a eles vai como veio.
  const comments = state.beforeComments
    ? state.comments
    : resolveComments(commentAnchorIdsOfJson(read), state.comments, new Set(state.commentsOutside))
  return {
    page: resolved.page,
    doc,
    styles: state.styles,
    ...(state.flattened ? { flattened: true } : {}),
    ...(state.beforeReferences ? { beforeReferences: true } : {}),
    ...(resolved.sections.length > 0 ? { sections: [...resolved.sections] } : {}),
    ...(state.beforeSections ? { beforeSections: true } : {}),
    ...(state.outsideBookmarks.length > 0 ? { outsideBookmarks: state.outsideBookmarks } : {}),
    ...(comments.length > 0 ? { comments } : {}),
    ...(state.beforeComments ? { beforeComments: true } : {}),
    ...(state.trackChanges === undefined ? {} : { trackChanges: state.trackChanges }),
    ...(state.beforeRevisions ? { beforeRevisions: true } : {}),
    ...(state.notes === undefined ? {} : { notes: state.notes }),
    ...(state.beforeNotes ? { beforeNotes: true } : {}),
    ...(state.beforeMath ? { beforeMath: true } : {}),
    ...(state.properties === undefined ? {} : { properties: state.properties }),
  }
}

/** O estado de um arquivo recém-aberto; sem arquivo, o da tela inicial. */
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
    // O aviso e a trava são deste arquivo, e morrem com ele.
    notice: null,
    savedLoss: null,
    readOnly: false,
  }
}
