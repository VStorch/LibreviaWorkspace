import type { StoreApi } from 'zustand'
import { AppError, type SerializedError } from '@shared/errors.js'
import { DiscardChoice } from '@shared/types.js'
import type { IpcResult } from '@shared/ipc.js'
import type { DocumentModel } from '@services/document/model.js'
import { serializeDocument } from '@services/document/serialize.js'
import { serializeWorkbook } from '@services/spreadsheet/serialize.js'
import type { DocumentSource, LoadedFile, WorkspaceState } from './types.js'

export type SetWorkspace = StoreApi<WorkspaceState>['setState']
export type GetWorkspace = StoreApi<WorkspaceState>['getState']

/**
 * O que todo grupo de ações precisa e nenhum deles é dono.
 *
 * Chamar o main tratando a falha, ler o que está na tela, esquecer o rascunho,
 * perguntar antes de descartar alterações e pôr um arquivo na tela: cinco
 * gestos que aparecem em abrir, salvar, recuperar e imprimir. Escritos uma vez
 * aqui, eles não têm como divergir entre um caminho e outro — e era em divergir
 * que estava o risco, porque quase todos mexem com o arquivo do usuário.
 */
export interface WorkspaceContext {
  /** Executa uma chamada de IPC e transforma a falha em erro exibível. */
  call: <T>(operation: () => Promise<IpcResult<T>>) => Promise<T | null>
  /** Modelo com o conteúdo que está na tela neste instante. */
  currentModel: () => DocumentModel
  /** O que está na tela, no formato interno — o mesmo que o rascunho guarda. */
  currentContent: () => string
  forgetDraft: () => Promise<void>
  /**
   * Portão de proteção contra perda de trabalho.
   *
   * Devolve `false` quando o usuário desistiu da ação — e nesse caso quem
   * chamou precisa parar sem alterar nada.
   */
  ensureChangesHandled: () => Promise<boolean>
  /** Põe um arquivo já interpretado na tela. */
  show: (loaded: LoadedFile) => void
  /** O editor montado, quando há um. */
  source: () => DocumentSource | null
  setSource: (source: DocumentSource | null) => void
}

export function toSerialized(cause: unknown): SerializedError {
  if (cause instanceof AppError) return cause.toSerialized()
  return { code: 'INTERNAL', message: 'Ocorreu um erro inesperado. A operação não foi concluída.' }
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
        error: { code: 'INTERNAL', message: 'A comunicação com o aplicativo falhou. Tente novamente.' },
      })
      return null
    } finally {
      set({ busy: false })
    }
  }

  function currentModel(): DocumentModel {
    const state = get()
    return { page: state.page, doc: documentSource?.readDoc() ?? state.initialDoc }
  }

  function currentContent(): string {
    const { workbook } = get()
    return workbook === null ? serializeDocument(currentModel()) : serializeWorkbook(workbook)
  }

  /**
   * Apaga o rascunho porque ele deixou de valer.
   *
   * Vale depois de gravar (o disco passou a ter a versão boa) e ao trocar de
   * arquivo (o rascunho é de outro trabalho). Falha aqui é engolida de
   * propósito: sobrar um rascunho velho custa um aviso a mais na próxima
   * abertura, e derrubar a gravação por causa disso custaria o arquivo.
   */
  async function forgetDraft(): Promise<void> {
    set({ autosaveBroken: false })
    await window.api.recovery.discard({}).catch(() => undefined)
  }

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

  function show({ file, model, workbook }: LoadedFile): void {
    // O rascunho é de outro trabalho a partir de agora. Vale inclusive para a
    // recuperação: o conteúdo já está na tela e marcado como não salvo, então o
    // relógio do autosave o grava de novo em segundos.
    void forgetDraft()

    set((state) => ({
      file,
      workbook,
      page: model.page,
      initialDoc: model.doc,
      generation: state.generation + 1,
      isDirty: false,
      error: null,
      // O aviso e a trava são deste arquivo, e morrem com ele: quem libera a
      // edição de um documento com comentários não está liberando o próximo, e
      // o inventário do anterior não fala deste. Quem abre um arquivo de
      // verdade preenche os dois logo em seguida, com o que o arquivo trouxe.
      notice: null,
      readOnly: false,
    }))
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
