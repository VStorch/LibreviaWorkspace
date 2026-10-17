import { z } from 'zod'
import { writeText, writeValues, type Range } from '@services/spreadsheet/edit.js'
import { recalculate } from '@services/spreadsheet/formula/recalc.js'
import { checkFormula } from '@services/spreadsheet/formula/validate.js'
import { getCell, type Cell, type Sheet, type WorkbookModel } from '@services/spreadsheet/model.js'
import { MAX_AI_FORMULA_LENGTH } from '@shared/limits.js'
import type { CellPosition } from './sheet-context.js'

export type FormulaAnswer =
  | { readonly kind: 'formula'; readonly formula: string }
  | { readonly kind: 'missing' }
  | { readonly kind: 'invalid'; readonly formula: string; readonly message: string }

export type FillValue = string | number | boolean | null

const FENCED = /```[^\n`]*\n?([\s\S]*?)```/
const INLINE_FORMULA = /`(=[^`\n]+)`/
const TRAILING_PERIOD = /\.$/

/** Models wrap the answer in a code block even when told not to. */
function unfenced(text: string): string {
  return FENCED.exec(text)?.[1] ?? text
}

const FORMULA_START = /(?:^|\s)=/g

/**
 * Where a formula could start in the line: at its beginning or after a space. In "Se A1 = 3, use
 * =SOMA(A1:A3)" the first `=` is prose.
 */
function formulaStarts(line: string): string[] {
  return [...line.matchAll(FORMULA_START)].map((match) =>
    line
      .slice(match.index + match[0].length - 1)
      .trim()
      .replace(TRAILING_PERIOD, ''),
  )
}

/**
 * "The formula is `=SOMA(A1:A3)`." or a line of its own. Of the candidates, the first the sheet
 * accepts; with none, the first, so the problem can be shown.
 */
function formulaIn(text: string): string | null {
  const inline = INLINE_FORMULA.exec(text)?.[1]
  if (inline !== undefined) return inline.trim()
  const candidates = unfenced(text).split('\n').flatMap(formulaStarts)
  return candidates.find((candidate) => checkFormula(candidate) === null) ?? candidates[0] ?? null
}

/** Only a formula that the sheet would take as typed: the same check as the formula bar. */
export function readFormulaAnswer(text: string): FormulaAnswer {
  const formula = formulaIn(text)
  if (formula === null || formula.length < 2 || formula.length > MAX_AI_FORMULA_LENGTH) {
    return { kind: 'missing' }
  }
  const problem = checkFormula(formula)
  return problem === null
    ? { kind: 'formula', formula }
    : { kind: 'invalid', formula, message: problem.message }
}

const fillValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()])

function jsonArrayIn(text: string): unknown {
  const body = unfenced(text)
  const start = body.indexOf('[')
  const end = body.lastIndexOf(']')
  if (start < 0 || end < start) return null
  try {
    return JSON.parse(body.slice(start, end + 1)) as unknown
  } catch {
    return null
  }
}

/** One value per row, in order, or `null` when the answer is not that. */
export function readFillAnswer(text: string, rowCount: number): readonly FillValue[] | null {
  const parsed = z.array(fillValueSchema).length(rowCount).safeParse(jsonArrayIn(text))
  return parsed.success ? parsed.data : null
}

/** Values, never formulas: a text that starts with `=` goes in as text. */
export function applyFill(sheet: Sheet, target: Range, values: readonly FillValue[]): Sheet {
  return writeValues(
    sheet,
    values.map((value, index) => ({ row: target.fromRow + index, column: target.fromColumn, value })),
  )
}

/** On a copy of the workbook: what the cell would show, without touching the open file. */
export function previewFormula(
  workbook: WorkbookModel,
  sheetIndex: number,
  at: CellPosition,
  formula: string,
): Cell | undefined {
  const sheets = [...workbook.sheets]
  const sheet = sheets[sheetIndex]
  if (sheet === undefined) return undefined
  sheets[sheetIndex] = writeText(sheet, at.row, at.column, formula)
  const calculated = recalculate({ ...workbook, sheets })
  return getCell(calculated.sheets[sheetIndex]!, at.row, at.column)
}
