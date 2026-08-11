import { create } from 'zustand'
import { AppError, type SerializedError } from '@shared/errors.js'
import type { IpcResult } from '@shared/ipc.js'
import { DiscardChoice, DocumentKind, PlainTextChoice, type RecentFile } from '@shared/types.js'
import {
  DEFAULT_PAGE_SETUP,
  createEmptyDocument,
  type DocumentModel,
  type DocumentNode,
  type PageSetup,
} from '@services/document/model.js'
import { documentToPlainText, hasRichFormatting, plainTextToDocument } from '@services/document/plain-text.js'
import { parseDocument, serializeDocument } from '@services/document/serialize.js'
import { defaultFileName, isPlainTextPath } from '@services/file/formats.js'

interface OpenFile {
  /** `null` enquanto o arquivo nunca foi gravado. */
  readonly path: string | null
  readonly name: string
  readonly kind: DocumentKind
}

interface WorkspaceState {
  file: OpenFile | null
  page: PageSetup
  /**
   * Documento com que o editor é inicializado. Não acompanha a digitação — o
   * conteúdo ao vivo mora no editor, e só é lido na hora de salvar.
   */
  initialDoc: DocumentNode
  /** Muda a cada novo/abrir; o editor é remontado, evitando estado residual. */
  generation: number
  isDirty: boolean
  /** Contagens exibidas na barra de status; alimentadas pelo editor. */
  stats: { characters: number; words: number }
  recents: readonly RecentFile[]
  error: SerializedError | null
  busy: boolean

  /** O editor se registra aqui para que salvar consiga ler o conteúdo atual. */
  registerDocumentSource: (read: (() => DocumentNode) | null) => void
  markDirty: () => void
  setStats: (stats: { characters: number; words: number }) => void
  setPage: (page: PageSetup) => void
  dismissError: () => void
  showError: (error: SerializedError) => void
  refreshRecents: () => Promise<void>

  newDocument: () => Promise<void>
  openViaDialog: () => Promise<void>
  openRecent: (path: string) => Promise<void>
  save: () => Promise<boolean>
  saveAs: () => Promise<boolean>
  closeFile: () => Promise<void>
  clearRecents: () => Promise<void>
}

function toSerialized(cause: unknown): SerializedError {
  if (cause instanceof AppError) return cause.toSerialized()
  return { code: 'INTERNAL', message: 'Ocorreu um erro inesperado. A operação não foi concluída.' }
}

export const useWorkspace = create<WorkspaceState>((set, get) => {
  let readLiveDoc: (() => DocumentNode) | null = null

  /** Executa uma chamada de IPC e transforma a falha em erro exibível. */
  async function call<T>(operation: () => Promise<IpcResult<T>>): Promise<T | null> {
    set({ busy: true })
    try {
      const result = await operation()
      if (result.ok) return result.data
      set({ error: result.error })
      return null
    } catch {
      set({
        error: { code: 'INTERNAL', message: 'A comunicação com o aplicativo falhou. Tente novamente.' },
      })
      return null
    } finally {
      set({ busy: false })
    }
  }

  /** Modelo com o conteúdo que está na tela neste instante. */
  function currentModel(): DocumentModel {
    const state = get()
    return { page: state.page, doc: readLiveDoc?.() ?? state.initialDoc }
  }

  /**
   * Portão de proteção contra perda de trabalho.
   *
   * Devolve `false` quando o usuário desistiu da ação — e nesse caso quem
   * chamou precisa parar sem alterar nada.
   */
  async function ensureChangesHandled(): Promise<boolean> {
    const state = get()
    if (!state.isDirty) return true

    const answer = await call(() =>
      window.api.dialog.confirmDiscard({ fileName: state.file?.name ?? 'Sem título' }),
    )
    if (answer === null) return false

    if (answer.choice === DiscardChoice.Cancel) return false
    if (answer.choice === DiscardChoice.Discard) return true
    return get().save()
  }

  function load(file: OpenFile, model: DocumentModel): void {
    set((state) => ({
      file,
      page: model.page,
      initialDoc: model.doc,
      generation: state.generation + 1,
      isDirty: false,
      error: null,
    }))
  }

  /** Interpreta o conteúdo lido do disco conforme a extensão do arquivo. */
  function decode(path: string, content: string): DocumentModel {
    if (isPlainTextPath(path)) {
      return { page: DEFAULT_PAGE_SETUP, doc: plainTextToDocument(content) }
    }
    return parseDocument(content)
  }

  async function openPath(load1: () => Promise<{ path: string; name: string; content: string } | null>) {
    const file = await load1()
    if (file === null) return

    try {
      load({ path: file.path, name: file.name, kind: DocumentKind.Document }, decode(file.path, file.content))
    } catch (cause) {
      set({ error: toSerialized(cause) })
    }
    await get().refreshRecents()
  }

  return {
    file: null,
    page: DEFAULT_PAGE_SETUP,
    initialDoc: createEmptyDocument().doc,
    generation: 0,
    isDirty: false,
    stats: { characters: 0, words: 0 },
    recents: [],
    error: null,
    busy: false,

    registerDocumentSource: (read) => {
      readLiveDoc = read
    },
    markDirty: () => {
      if (!get().isDirty) set({ isDirty: true })
    },
    setStats: (stats) => {
      const current = get().stats
      if (current.characters !== stats.characters || current.words !== stats.words) set({ stats })
    },
    setPage: (page) => set({ page, isDirty: true }),
    dismissError: () => set({ error: null }),
    showError: (error) => set({ error }),

    refreshRecents: async () => {
      const data = await call(() => window.api.recent.list({}))
      if (data !== null) set({ recents: data.files })
    },

    newDocument: async () => {
      if (!(await ensureChangesHandled())) return
      const empty = createEmptyDocument()
      load({ path: null, name: defaultFileName(DocumentKind.Document), kind: DocumentKind.Document }, empty)
    },

    openViaDialog: async () => {
      if (!(await ensureChangesHandled())) return
      await openPath(async () => {
        const data = await call(() => window.api.file.open({}))
        return data === null || data.canceled ? null : data.file
      })
    },

    openRecent: async (path) => {
      if (!(await ensureChangesHandled())) return
      await openPath(async () => {
        const data = await call(() => window.api.file.openRecent({ path }))
        if (data === null) {
          // O arquivo pode ter sumido; a lista precisa refletir isso.
          await get().refreshRecents()
          return null
        }
        return data.file
      })
    },

    save: async () => {
      const state = get()
      if (state.file === null) return false
      // Arquivo que nunca foi gravado não tem destino: vira "salvar como".
      if (state.file.path === null) return get().saveAs()

      const model = currentModel()
      let content: string

      if (isPlainTextPath(state.file.path)) {
        if (hasRichFormatting(model.doc)) {
          // Salvar em .txt descartaria formatação. Perguntar antes é a regra
          // do projeto: nada se perde em silêncio.
          const answer = await call(() =>
            window.api.dialog.confirmPlainText({ fileName: state.file?.name ?? 'documento' }),
          )
          if (answer === null || answer.choice === PlainTextChoice.Cancel) return false
          if (answer.choice === PlainTextChoice.SaveAsDocument) return get().saveAs()
        }
        content = documentToPlainText(model.doc)
      } else {
        content = serializeDocument(model)
      }

      const data = await call(() => window.api.file.save({ path: state.file?.path ?? '', content }))
      if (data === null) return false

      set({ isDirty: false, file: { ...state.file, name: data.name } })
      await get().refreshRecents()
      return true
    },

    saveAs: async () => {
      const state = get()
      if (state.file === null) return false

      const model = currentModel()
      const chosen = await call(() =>
        window.api.file.chooseSavePath({
          suggestedName: state.file?.name ?? defaultFileName(DocumentKind.Document),
        }),
      )
      if (chosen === null || chosen.canceled) return false

      // O aviso vem antes da gravação: se o usuário desistir aqui, nenhum byte
      // foi escrito e o destino continua como estava.
      if (isPlainTextPath(chosen.path) && hasRichFormatting(model.doc)) {
        const answer = await call(() => window.api.dialog.confirmPlainText({ fileName: chosen.name }))
        if (answer === null || answer.choice === PlainTextChoice.Cancel) return false
        // Quis preservar a formatação: escolhe outro destino.
        if (answer.choice === PlainTextChoice.SaveAsDocument) return get().saveAs()
      }

      const content = isPlainTextPath(chosen.path) ? documentToPlainText(model.doc) : serializeDocument(model)

      const data = await call(() => window.api.file.save({ path: chosen.path, content }))
      if (data === null) return false

      set({ file: { path: chosen.path, name: data.name, kind: state.file.kind }, isDirty: false })
      await get().refreshRecents()
      return true
    },

    closeFile: async () => {
      if (!(await ensureChangesHandled())) return
      const empty = createEmptyDocument()
      set((state) => ({
        file: null,
        page: empty.page,
        initialDoc: empty.doc,
        generation: state.generation + 1,
        isDirty: false,
        error: null,
      }))
      await get().refreshRecents()
    },

    clearRecents: async () => {
      const data = await call(() => window.api.recent.clear({}))
      if (data !== null) set({ recents: data.files })
    },
  }
})
