import type { WorkbookModel } from './model.js'

export interface WorkbookStatistics {
  readonly sheets: number
  /** With a value or a formula: a cell with only formatting does not count, as in Excel. */
  readonly filledCells: number
  readonly formulas: number
}

export function workbookStatistics(workbook: WorkbookModel): WorkbookStatistics {
  let filledCells = 0
  let formulas = 0
  for (const sheet of workbook.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      if (cell.formula !== undefined) formulas++
      if (cell.formula !== undefined || (cell.value !== undefined && cell.value !== '')) filledCells++
    }
  }
  return { sheets: workbook.sheets.length, filledCells, formulas }
}
