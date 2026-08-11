import { create } from 'zustand'
import type { SerializedError } from '@shared/errors.js'
import type { IpcResult } from '@shared/ipc.js'
import { DiscardChoice, DocumentKind, type RecentFile } from '@shared/types.js'
import { defaultFileName } from '@services/file/formats.js'

interface OpenFile {
  /** `null` enquanto o arquivo nunca foi gravado. */
  readonly path: string | null
  readonly name: string
  readonly kind: DocumentKind
}

interface WorkspaceState {
  file: OpenFile | null
  content: string
  /** Conteúdo da última gravação. A comparação com `content` define o "sujo". */
  savedContent: string
  recents: readonly RecentFile[]
  error: SerializedError | null
  busy: boolean

  setContent: (content: string) => void
  dismissError: () => void
  refreshRecents: () => Promise<void>

  newDocument: () => Promise<void>
  openViaDialog: () => Promise<void>
  openRecent: (path: string) => Promise<void>
  save: () => Promise<boolean>
  saveAs: () => Promise<boolean>
  closeFile: () => Promise<void>
  clearRecents: () => Promise<void>
}

export const isDirty = (state: WorkspaceState): boolean =>
  state.file !== null && state.content !== state.savedContent

export const useWorkspace = create<WorkspaceState>((set, get) => {
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
        error: {
          code: 'INTERNAL',
          message: 'A comunicação com o aplicativo falhou. Tente novamente.',
        },
      })
      return null
    } finally {
      set({ busy: false })
    }
  }

  /**
   * Portão de proteção contra perda de trabalho.
   *
   * Devolve `false` quando o usuário desistiu da ação — e nesse caso quem
   * chamou precisa parar sem alterar nada.
   */
  async function ensureChangesHandled(): Promise<boolean> {
    const state = get()
    if (!isDirty(state)) return true

    const answer = await call(() =>
      window.api.dialog.confirmDiscard({ fileName: state.file?.name ?? 'Sem título' }),
    )
    if (answer === null) return false

    if (answer.choice === DiscardChoice.Cancel) return false
    if (answer.choice === DiscardChoice.Discard) return true
    return get().save()
  }

  function adopt(path: string, name: string, content: string, kind: DocumentKind): void {
    set({ file: { path, name, kind }, content, savedContent: content, error: null })
  }

  return {
    file: null,
    content: '',
    savedContent: '',
    recents: [],
    error: null,
    busy: false,

    setContent: (content) => set({ content }),
    dismissError: () => set({ error: null }),

    refreshRecents: async () => {
      const data = await call(() => window.api.recent.list({}))
      if (data !== null) set({ recents: data.files })
    },

    newDocument: async () => {
      if (!(await ensureChangesHandled())) return
      set({
        file: { path: null, name: defaultFileName(DocumentKind.Document), kind: DocumentKind.Document },
        content: '',
        savedContent: '',
        error: null,
      })
    },

    openViaDialog: async () => {
      if (!(await ensureChangesHandled())) return
      const data = await call(() => window.api.file.open({}))
      if (data === null || data.canceled) return

      adopt(data.file.path, data.file.name, data.file.content, data.file.kind)
      await get().refreshRecents()
    },

    openRecent: async (path) => {
      if (!(await ensureChangesHandled())) return
      const data = await call(() => window.api.file.openRecent({ path }))
      if (data === null) {
        // O arquivo pode ter sumido; a lista precisa refletir isso.
        await get().refreshRecents()
        return
      }

      adopt(data.file.path, data.file.name, data.file.content, data.file.kind)
      await get().refreshRecents()
    },

    save: async () => {
      const state = get()
      if (state.file === null) return false
      // Arquivo que nunca foi gravado não tem destino: vira "salvar como".
      if (state.file.path === null) return get().saveAs()

      const data = await call(() =>
        window.api.file.save({ path: state.file?.path ?? '', content: state.content }),
      )
      if (data === null) return false

      set({ savedContent: state.content, file: { ...state.file, name: data.name } })
      await get().refreshRecents()
      return true
    },

    saveAs: async () => {
      const state = get()
      if (state.file === null) return false

      const data = await call(() =>
        window.api.file.saveAs({
          suggestedName: state.file?.name ?? defaultFileName(DocumentKind.Document),
          kind: state.file?.kind ?? DocumentKind.Document,
          content: state.content,
        }),
      )
      if (data === null || data.canceled) return false

      set({
        file: { path: data.path, name: data.name, kind: state.file.kind },
        savedContent: state.content,
      })
      await get().refreshRecents()
      return true
    },

    closeFile: async () => {
      if (!(await ensureChangesHandled())) return
      set({ file: null, content: '', savedContent: '', error: null })
      await get().refreshRecents()
    },

    clearRecents: async () => {
      const data = await call(() => window.api.recent.clear({}))
      if (data !== null) set({ recents: data.files })
    },
  }
})
