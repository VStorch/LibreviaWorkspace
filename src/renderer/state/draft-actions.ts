import type { SerializedError } from '@shared/errors.js'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import { createEmptyDocument } from '@services/document/model.js'
import { parseDocument } from '@services/document/serialize.js'
import { parseWorkbook } from '@services/spreadsheet/serialize.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import { toSerialized, type GetWorkspace, type SetWorkspace, type WorkspaceContext } from './context.js'
import type { LoadedFile, WorkspaceState } from './types.js'

/** O rascunho como ele volta do disco: o resumo mais o conteúdo. */
type RecoveredDraft = DraftSummary & { readonly content: string }

type DraftActions = Pick<WorkspaceState, 'autosave' | 'checkRecovery' | 'recoverDraft' | 'dismissDraft'>

export function createDraftActions(
  set: SetWorkspace,
  get: GetWorkspace,
  ctx: WorkspaceContext,
): DraftActions {
  return {
    autosave: async () => {
      const state = get()
      // Sem arquivo aberto não há o que guardar; sem alteração, o rascunho já
      // está em dia. E enquanto houver rascunho esperando decisão, escrever
      // apagaria o trabalho que ele existe para devolver.
      if (state.file === null || !state.isDirty) return
      if (state.pendingDraft !== null || state.autosaveBroken) return

      // Fora do `call`: o autosave não pode piscar o indicador de ocupado nem
      // roubar a vez de uma operação que o usuário pediu.
      const result = await window.api.file.autosave({
        path: state.file.path,
        name: state.file.name,
        kind: state.file.kind,
        content: ctx.currentContent(),
      })

      if (!result.ok) set(autosaveBrokenState(result.error))
    },

    checkRecovery: async () => {
      const result = await window.api.recovery.peek({})
      if (result.ok && result.data.draft !== null) set({ pendingDraft: result.data.draft })
    },

    recoverDraft: async () => {
      const data = await ctx.call(() => window.api.recovery.restore({}))
      set({ pendingDraft: null })
      if (data === null || data.draft === null) return

      try {
        ctx.show(interpret(data.draft))
        // Recuperado é, por definição, diferente do que está no disco: marcar
        // como salvo faria o usuário fechar a janela e perder tudo de novo.
        set({ isDirty: true })
      } catch (cause) {
        set({ error: toSerialized(cause) })
      }
    },

    dismissDraft: async () => {
      set({ pendingDraft: null })
      await window.api.recovery.discard({})
    },
  }
}

/**
 * O rascunho guarda sempre o formato interno, mesmo o de um arquivo `.txt`: ele
 * é o que estava na tela, e não o que o disco tem. Por isso a leitura aqui não
 * olha a extensão — quem olha é a abertura de um arquivo de verdade.
 */
function interpret(draft: RecoveredDraft): LoadedFile {
  const file = { path: draft.path, name: draft.name, kind: draft.kind }

  if (draft.kind === DocumentKind.Spreadsheet) {
    return { file, model: createEmptyDocument(), workbook: recalculate(parseWorkbook(draft.content)) }
  }

  return { file, model: parseDocument(draft.content), workbook: null }
}

/**
 * O autosave falhou: avisa uma vez e para de tentar.
 *
 * Insistir a cada oito segundos numa gravação que não vai dar certo encheria a
 * tela, e falhar em silêncio deixaria o usuário confiando numa proteção que não
 * existe mais.
 *
 * A frase é reescrita porque a original fala do arquivo do usuário ("o conteúdo
 * original foi preservado"), e aqui nenhum arquivo dele foi tocado. O que
 * quebrou foi a rede de proteção, e é isso que precisa ser dito — junto com o
 * que fazer a respeito.
 */
function autosaveBrokenState(cause: SerializedError): Partial<WorkspaceState> {
  return {
    autosaveBroken: true,
    error: {
      code: cause.code,
      message:
        'A gravação automática de segurança parou de funcionar. Seu arquivo não foi alterado, ' +
        'mas salve o trabalho manualmente: uma queda agora custaria o que não foi salvo.',
      detail: cause.message,
    },
  }
}
