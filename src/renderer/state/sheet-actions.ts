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
    /**
     * Toda mudança de célula passa por aqui, e por isso o recálculo mora aqui.
     *
     * Um só ponto de recálculo é o que garante que nenhum caminho de edição
     * deixe um valor velho na tela: quem esquecer de recalcular teria uma
     * planilha que só mostra o resultado certo na próxima vez que for tocada.
     */
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
      // Trocar de aba não suja o arquivo: é navegação, não edição.
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

      // Nome repetido quebraria a referência entre abas — melhor recusar agora,
      // em silêncio, do que aceitar e falhar depois.
      if (isNameTaken(workbook, trimmed, index)) return

      // Renomear reescreve as fórmulas que citam a aba: sem isso o gesto, que o
      // usuário considera cosmético, viraria #REF! em toda planilha que a usa.
      set({ workbook: recalculate(renameSheetIn(workbook, index, trimmed)), isDirty: true })
    },

    removeSheet: (index) => {
      const { workbook } = get()
      // Uma pasta sem planilha nenhuma não é estado válido do modelo.
      if (workbook === null || workbook.sheets.length <= 1) return

      const sheets = workbook.sheets.filter((_, at) => at !== index)
      const activeSheet = Math.min(workbook.activeSheet, sheets.length - 1)
      set({ workbook: { sheets, activeSheet }, isDirty: true })
    },
  }
}
