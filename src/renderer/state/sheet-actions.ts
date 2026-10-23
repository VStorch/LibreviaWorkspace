import { createSheet, type WorkbookModel } from '@services/spreadsheet/model.js'
import type { StoreApi } from 'zustand'
import {
  changedCells,
  emptyHistory,
  recordEdit,
  redoStep,
  undoStep,
  type History,
  type Step,
} from '@services/spreadsheet/history.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import {
  applyStructuralChange,
  isNameTaken,
  nextSheetName,
  renameSheet as renameSheetIn,
} from '@services/spreadsheet/structure.js'
import { currentPreferences } from './preferences.js'
import type { GetWorkspace, SetWorkspace } from './context.js'
import type { WorkspaceState } from './types.js'

type SheetActions = Pick<
  WorkspaceState,
  | 'updateSheet'
  | 'changeStructure'
  | 'selectSheet'
  | 'addSheet'
  | 'renameSheet'
  | 'removeSheet'
  | 'undoSheet'
  | 'redoSheet'
>

type Move = (
  history: History<WorkbookModel>,
  current: WorkbookModel,
  weigh: typeof changedCells,
) => Step<WorkbookModel> | null

type Subscribe = StoreApi<WorkspaceState>['subscribe']

/** What was last loaded or saved: undoing back to it leaves nothing to save. */
interface SavedPoint {
  readonly workbook: WorkbookModel | null
  readonly properties: WorkspaceState['properties']
}

interface Kept {
  readonly generation: number
  readonly history: History<WorkbookModel>
  readonly saved: SavedPoint | null
}

/**
 * The history belongs to the open file: another one drops it at once, so the old workbook is not
 * held in memory.
 */
function createWorkbookHistory(set: SetWorkspace, get: GetWorkspace, subscribe: Subscribe) {
  let kept: Kept = { generation: -1, history: emptyHistory(), saved: null }

  const follow = (state: WorkspaceState): void => {
    if (state.generation !== kept.generation) {
      kept = { generation: state.generation, history: emptyHistory(), saved: null }
    }
    if (!state.isDirty) {
      kept = { ...kept, saved: { workbook: state.workbook, properties: state.properties } }
    }
  }
  // Not read now: the store is still being created.
  subscribe(follow)

  const isSaved = (workbook: WorkbookModel): boolean =>
    kept.saved?.workbook === workbook && kept.saved.properties === get().properties

  const commit = (previous: WorkbookModel, next: WorkbookModel): void => {
    kept = { ...kept, history: recordEdit(kept.history, previous, next, changedCells) }
    set({ workbook: next, isDirty: true })
  }

  const travel = (move: Move): void => {
    const { workbook } = get()
    if (workbook === null) return
    const step = move(kept.history, workbook, changedCells)
    if (step === null) return
    kept = { ...kept, history: step.history }
    set({ workbook: step.state, isDirty: !isSaved(step.state) })
  }

  return { commit, undo: () => travel(undoStep), redo: () => travel(redoStep) }
}

export function createSheetActions(set: SetWorkspace, get: GetWorkspace, subscribe: Subscribe): SheetActions {
  const history = createWorkbookHistory(set, get, subscribe)
  return {
    /** A single recalculation point: no editing path leaves a stale value on screen. */
    updateSheet: (sheet) => {
      const { workbook } = get()
      if (workbook === null || workbook.sheets[workbook.activeSheet] === sheet) return

      const sheets = [...workbook.sheets]
      sheets[workbook.activeSheet] = sheet
      history.commit(workbook, recalculate({ ...workbook, sheets }))
    },

    changeStructure: (change) => {
      const { workbook } = get()
      if (workbook === null) return

      const changed = applyStructuralChange(workbook, workbook.activeSheet, change)
      if (changed === workbook) return

      history.commit(workbook, recalculate(changed))
    },

    selectSheet: (index) => {
      const { workbook } = get()
      if (workbook === null || index < 0 || index >= workbook.sheets.length) return
      // Navigation, not editing.
      set({ workbook: { ...workbook, activeSheet: index } })
    },

    addSheet: () => {
      const { workbook } = get()
      if (workbook === null) return

      const sheets = [...workbook.sheets, createSheet(nextSheetName(workbook, currentPreferences().language))]
      history.commit(workbook, { sheets, activeSheet: sheets.length - 1 })
    },

    renameSheet: (index, name) => {
      const { workbook } = get()
      if (workbook === null) return

      const trimmed = name.trim()
      const sheet = workbook.sheets[index]
      if (sheet === undefined || trimmed.length === 0) return

      // A repeated name would break references between tabs.
      if (isNameTaken(workbook, trimmed, index)) return

      history.commit(workbook, recalculate(renameSheetIn(workbook, index, trimmed)))
    },

    removeSheet: (index) => {
      const { workbook } = get()
      if (workbook === null || workbook.sheets.length <= 1) return

      const sheets = workbook.sheets.filter((_, at) => at !== index)
      const activeSheet = Math.min(workbook.activeSheet, sheets.length - 1)
      history.commit(workbook, { sheets, activeSheet })
    },

    undoSheet: history.undo,
    redoSheet: history.redo,
  }
}
