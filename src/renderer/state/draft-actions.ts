import type { SerializedError } from '@shared/errors.js'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import { createEmptyDocument } from '@services/document/model.js'
import { parseDocument } from '@services/document/serialize.js'
import { parseWorkbook } from '@services/spreadsheet/serialize.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import { toSerialized, type GetWorkspace, type SetWorkspace, type WorkspaceContext } from './context.js'
import { t } from '../i18n.js'
import type { LoadedFile, WorkspaceState } from './types.js'

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
      // Com um rascunho esperando decisão, escrever apagaria o trabalho que ele devolve.
      if (state.file === null || !state.isDirty) return
      if (state.pendingDraft !== null || state.autosaveBroken) return

      // Fora do `call`: o autosave não pisca o indicador de ocupado.
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
        // Recuperado difere do disco: marcar como salvo faria fechar e perder tudo de novo.
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

/** Sempre em formato interno, mesmo de um `.txt`: é o que estava na tela. */
function interpret(draft: RecoveredDraft): LoadedFile {
  const file = { path: draft.path, name: draft.name, kind: draft.kind }

  if (draft.kind === DocumentKind.Spreadsheet) {
    return { file, model: createEmptyDocument(), workbook: recalculate(parseWorkbook(draft.content)) }
  }

  return { file, model: parseDocument(draft.content), workbook: null }
}

/**
 * Avisa uma vez e para de tentar. A frase é própria, porque a do erro fala do
 * arquivo do usuário, e aqui o que quebrou foi a rede de proteção.
 */
function autosaveBrokenState(cause: SerializedError): Partial<WorkspaceState> {
  return {
    autosaveBroken: true,
    error: {
      code: cause.code,
      message: t('shell.draft.autosaveBroken'),
      detail: cause.message,
    },
  }
}
