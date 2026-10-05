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
      // With a draft awaiting a decision, writing would erase the work it brings back.
      if (state.file === null || !state.isDirty) return
      if (state.pendingDraft !== null || state.autosaveBroken) return

      // Outside `call`: autosave does not flash the busy indicator.
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
        // Recovered content differs from disk: marking it saved would let a close lose everything
        // again.
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

/** Always in the internal format, even from a `.txt`: it is what was on screen. */
function interpret(draft: RecoveredDraft): LoadedFile {
  const file = { path: draft.path, name: draft.name, kind: draft.kind }

  if (draft.kind === DocumentKind.Spreadsheet) {
    return { file, model: createEmptyDocument(), workbook: recalculate(parseWorkbook(draft.content)) }
  }

  return { file, model: parseDocument(draft.content), workbook: null }
}

/**
 * Warns once and stops trying. Its own sentence, because the error one speaks of the user's file,
 * and what broke here is the safety net.
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
