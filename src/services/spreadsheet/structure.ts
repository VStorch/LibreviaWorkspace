/**
 * On the **workbook**, not the sheet: a row inserted in "Dados" changes `=Dados!A5` written in
 * "Resumo".
 */

import { adjustForColumns, adjustForRows, renameSheetInFormula } from './formula/adjust.js'
import {
  deleteColumns as deleteColumnsIn,
  deleteRows as deleteRowsIn,
  insertColumns as insertColumnsIn,
  insertRows as insertRowsIn,
} from './edit.js'
import type { Cell, Sheet, WorkbookModel } from './model.js'

/** As the UI describes it. */
export type StructuralChange =
  | { readonly kind: 'insertRows'; readonly at: number; readonly count: number }
  | { readonly kind: 'deleteRows'; readonly at: number; readonly count: number }
  | { readonly kind: 'insertColumns'; readonly at: number; readonly count: number }
  | { readonly kind: 'deleteColumns'; readonly at: number; readonly count: number }

export function applyStructuralChange(
  workbook: WorkbookModel,
  sheetIndex: number,
  change: StructuralChange,
): WorkbookModel {
  const target = workbook.sheets[sheetIndex]
  if (target === undefined) return workbook

  const shifted = shift(target, change)
  if (shifted === target) return workbook

  const axis = change.kind === 'insertRows' || change.kind === 'deleteRows' ? adjustForRows : adjustForColumns
  const delta = change.kind.startsWith('insert') ? change.count : -change.count

  const sheets = workbook.sheets.map((sheet, index) => {
    const base = index === sheetIndex ? shifted : sheet
    return rewriteFormulas(base, (formula) =>
      axis(formula, change.at, delta, { sheet: target.name, own: index === sheetIndex }),
    )
  })

  return { ...workbook, sheets }
}

function shift(sheet: Sheet, change: StructuralChange): Sheet {
  switch (change.kind) {
    case 'insertRows':
      return insertRowsIn(sheet, change.at, change.count)
    case 'deleteRows':
      return deleteRowsIn(sheet, change.at, change.count)
    case 'insertColumns':
      return insertColumnsIn(sheet, change.at, change.count)
    case 'deleteColumns':
      return deleteColumnsIn(sheet, change.at, change.count)
  }
}

/** Without the rewrite, renaming would turn every formula citing the tab into `#REF!`. */
export function renameSheet(workbook: WorkbookModel, sheetIndex: number, name: string): WorkbookModel {
  const target = workbook.sheets[sheetIndex]
  if (target === undefined || target.name === name) return workbook

  const sheets = workbook.sheets.map((sheet, index) => {
    const renamed = index === sheetIndex ? { ...sheet, name } : sheet
    return rewriteFormulas(renamed, (formula) => renameSheetInFormula(formula, target.name, name))
  })

  return { ...workbook, sheets }
}

/** Skips used names: someone who deleted Planilha2 and created another would have two. */
export function nextSheetName(workbook: WorkbookModel): string {
  const used = new Set(workbook.sheets.map((sheet) => sheet.name))

  let index = workbook.sheets.length + 1
  while (used.has(`Planilha${index}`)) index += 1
  return `Planilha${index}`
}

/** Is there already a tab with this name, not counting the one being renamed? */
export function isNameTaken(workbook: WorkbookModel, name: string, exceptIndex: number): boolean {
  return workbook.sheets.some((sheet, index) => index !== exceptIndex && sheet.name === name)
}

/** A sheet without any formula comes back as the same object. */
function rewriteFormulas(sheet: Sheet, rewrite: (formula: string) => string): Sheet {
  let cells: Record<string, Cell> | null = null

  for (const [ref, cell] of Object.entries(sheet.cells)) {
    if (cell.formula === undefined) continue

    const formula = rewrite(cell.formula)
    if (formula === cell.formula) continue

    cells ??= { ...sheet.cells }
    cells[ref] = { ...cell, formula }
  }

  return cells === null ? sheet : { ...sheet, cells }
}
