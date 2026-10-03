import { createSheet } from '@services/spreadsheet/model.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import {
  applyStructuralChange,
  isNameTaken,
  nextSheetName,
  renameSheet as renameSheetIn,
} from '@services/spreadsheet/structure.js'
import type { GetWorkspace, SetWorkspace } from './context.js'
import type { WorkspaceState } from './types.js'

type SheetActions = Pick<
  WorkspaceState,
  'updateSheet' | 'changeStructure' | 'selectSheet' | 'addSheet' | 'renameSheet' | 'removeSheet'
>

export function createSheetActions(set: SetWorkspace, get: GetWorkspace): SheetActions {
  return {
    /** Um só ponto de recálculo: nenhum caminho de edição deixa valor velho na tela. */
    updateSheet: (sheet) => {
      const { workbook } = get()
      if (workbook === null) return

      const sheets = [...workbook.sheets]
      sheets[workbook.activeSheet] = sheet
      set({ workbook: recalculate({ ...workbook, sheets }), isDirty: true })
    },

    changeStructure: (change) => {
      const { workbook } = get()
      if (workbook === null) return

      const changed = applyStructuralChange(workbook, workbook.activeSheet, change)
      if (changed === workbook) return

      set({ workbook: recalculate(changed), isDirty: true })
    },

    selectSheet: (index) => {
      const { workbook } = get()
      if (workbook === null || index < 0 || index >= workbook.sheets.length) return
      // Navegação, e não edição.
      set({ workbook: { ...workbook, activeSheet: index } })
    },

    addSheet: () => {
      const { workbook } = get()
      if (workbook === null) return

      const sheets = [...workbook.sheets, createSheet(nextSheetName(workbook))]
      set({ workbook: { sheets, activeSheet: sheets.length - 1 }, isDirty: true })
    },

    renameSheet: (index, name) => {
      const { workbook } = get()
      if (workbook === null) return

      const trimmed = name.trim()
      const sheet = workbook.sheets[index]
      if (sheet === undefined || trimmed.length === 0) return

      // Nome repetido quebraria a referência entre abas.
      if (isNameTaken(workbook, trimmed, index)) return

      set({ workbook: recalculate(renameSheetIn(workbook, index, trimmed)), isDirty: true })
    },

    removeSheet: (index) => {
      const { workbook } = get()
      if (workbook === null || workbook.sheets.length <= 1) return

      const sheets = workbook.sheets.filter((_, at) => at !== index)
      const activeSheet = Math.min(workbook.activeSheet, sheets.length - 1)
      set({ workbook: { sheets, activeSheet }, isDirty: true })
    },
  }
}
