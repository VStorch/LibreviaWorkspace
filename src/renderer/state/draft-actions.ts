import type { SerializedError } from '@shared/errors.js'
import { DocumentKind } from '@shared/types.js'
import { createEmptyDocument } from '@services/document/model.js'
import { parseDocument } from '@services/document/serialize.js'
import { parseWorkbook } from '@services/spreadsheet/serialize.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import { toSerialized, type GetWorkspace, type SetWorkspace, type WorkspaceContext } from './context.js'
import { t } from '../i18n.js'
import type { LoadedFile, RecoveredDraft, WorkspaceState } from './types.js'

type DraftActions = Pick<WorkspaceState, 'autosave' | 'restoreDraft'>

export function createDraftActions(
  set: SetWorkspace,
  get: GetWorkspace,
  ctx: WorkspaceContext,
): DraftActions {
  /**
   * Without an editor, the content lives in the state: the same state means the same draft. A
   * hidden tab would otherwise write it again every few seconds.
   */
  let written: WorkspaceState | null = null

  return {
    autosave: async () => {
      const state = get()
      if (state.file === null || !state.isDirty) return
      if (ctx.source() === null && state === written) return
      // A PDF's only change is its properties, which have no draft: closing asks to save them.
      if (state.file.kind === DocumentKind.Pdf) return
      if (state.autosaveBroken) return

      // Outside `call`: autosave does not flash the busy indicator.
      const result = await window.api.file.autosave({
        draftId: state.draftId,
        path: state.file.path,
        name: state.file.name,
        kind: state.file.kind,
        content: ctx.currentContent(),
      })

      if (result.ok) written = state
      else set(autosaveBrokenState(result.error))
    },

    restoreDraft: (draft) => {
      set({ draftId: draft.draftId })
      try {
        ctx.show(interpret(draft))
        // Recovered content differs from disk: marking it saved would let a close lose everything
        // again.
        set({ isDirty: true, textEdited: true })
        return true
      } catch (cause) {
        set({ error: toSerialized(cause) })
        return false
      }
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
