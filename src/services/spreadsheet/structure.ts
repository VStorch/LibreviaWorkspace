/**
 * Inserir e excluir linhas e colunas na pasta inteira.
 *
 * As operações de `edit.ts` mexem numa planilha só, porque é só de células que
 * elas tratam. Fórmula é outra história: uma linha inserida em "Dados" muda o
 * significado de `=Dados!A5` escrita em "Resumo". Por isso a operação estrutural
 * é da **pasta**, e não da planilha — e é por isso que ela vive aqui, e não lá.
 */

import { adjustForColumns, adjustForRows, renameSheetInFormula } from './formula/adjust.js'
import {
  deleteColumns as deleteColumnsIn,
  deleteRows as deleteRowsIn,
  insertColumns as insertColumnsIn,
  insertRows as insertRowsIn,
} from './edit.js'
import type { Cell, Sheet, WorkbookModel } from './model.js'

/** Uma mudança de estrutura, como a interface a descreve. */
export type StructuralChange =
  | { readonly kind: 'insertRows'; readonly at: number; readonly count: number }
  | { readonly kind: 'deleteRows'; readonly at: number; readonly count: number }
  | { readonly kind: 'insertColumns'; readonly at: number; readonly count: number }
  | { readonly kind: 'deleteColumns'; readonly at: number; readonly count: number }

/**
 * Aplica a mudança e reescreve as referências de todas as planilhas.
 *
 * Sem a reescrita, inserir uma linha moveria os dados e deixaria `SOMA(A1:A3)`
 * apontando para onde eles não estão mais — um total errado, sem aviso.
 */
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

/**
 * Renomeia uma planilha e conserta as fórmulas que a citam.
 *
 * Renomear é um gesto que o usuário considera cosmético; sem esta reescrita ele
 * transformaria em `#REF!` toda fórmula que apontava para a aba.
 */
export function renameSheet(workbook: WorkbookModel, sheetIndex: number, name: string): WorkbookModel {
  const target = workbook.sheets[sheetIndex]
  if (target === undefined || target.name === name) return workbook

  const sheets = workbook.sheets.map((sheet, index) => {
    const renamed = index === sheetIndex ? { ...sheet, name } : sheet
    return rewriteFormulas(renamed, (formula) => renameSheetInFormula(formula, target.name, name))
  })

  return { ...workbook, sheets }
}

/**
 * O nome da próxima aba: "Planilha2", "Planilha3"…, pulando os já usados.
 *
 * Nome repetido quebraria a referência entre abas — `=Planilha2!A1` deixaria de
 * ter destino único —, e é por isso que a contagem não é simplesmente o total
 * de abas mais um: quem apagou a Planilha2 e criou outra teria duas.
 */
export function nextSheetName(workbook: WorkbookModel): string {
  const used = new Set(workbook.sheets.map((sheet) => sheet.name))

  let index = workbook.sheets.length + 1
  while (used.has(`Planilha${index}`)) index += 1
  return `Planilha${index}`
}

/** A aba já existe com este nome, sem contar a que está sendo renomeada? */
export function isNameTaken(workbook: WorkbookModel, name: string, exceptIndex: number): boolean {
  return workbook.sheets.some((sheet, index) => index !== exceptIndex && sheet.name === name)
}

/** Planilha sem fórmula nenhuma volta como o mesmo objeto. */
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
