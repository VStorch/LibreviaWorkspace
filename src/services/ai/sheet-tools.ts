import { z } from 'zod'
import { describeRange, parseRange, rangeContains, type Range } from '@services/spreadsheet/edit.js'
import { formatCell } from '@services/spreadsheet/format.js'
import { quoteSheet } from '@services/spreadsheet/formula/references.js'
import { checkFormula } from '@services/spreadsheet/formula/validate.js'
import { cellRef, parseRef, type Sheet, type WorkbookModel } from '@services/spreadsheet/model.js'
import {
  MAX_AI_FORMULA_LENGTH,
  MAX_AI_READ_CELLS,
  MAX_AI_REASON_LENGTH,
  MAX_AI_SEARCH_QUERY_LENGTH,
  MAX_AI_SEARCH_RESULTS,
  MAX_AI_WRITE_CELLS,
} from '@shared/limits.js'
import type { AiToolCall, AiToolDefinition } from './provider.js'
import { byPosition, cellLine, usedCells, type UsedCell } from './sheet-context.js'
import { defineTool as defineAnyTool, failure, success, toolbox, type ToolOutcome } from './tools.js'

/** The workbook on screen and the selection on its active sheet. */
export interface SheetToolContext {
  readonly workbook: WorkbookModel
  readonly selection: Range
}

/** Cells to write as typed, from the top left corner: a text that starts with `=` is a formula. */
export interface SheetProposal {
  readonly sheet: string
  readonly row: number
  readonly column: number
  readonly rows: readonly (readonly string[])[]
  readonly reason?: string
}

type SheetOutcome = ToolOutcome<SheetProposal>

export const SheetToolName = {
  ListSheets: 'list_sheets',
  ReadCells: 'read_cells',
  FindCells: 'find_cells',
  WriteCells: 'write_cells',
} as const

const DEFAULT_FIND_RESULTS = 10
const MAX_REF_LENGTH = 16
const RANGE_FORMAT = 'The range reads like "A1" or "A1:D20".'

function defineTool<S extends z.ZodType>(
  name: string,
  description: string,
  schema: S,
  run: (args: z.infer<S>, context: SheetToolContext) => SheetOutcome,
) {
  return defineAnyTool(name, description, schema, run)
}

const sheetName = z
  .string()
  .min(1)
  .max(MAX_AI_SEARCH_QUERY_LENGTH)
  .optional()
  .describe('Name of the sheet; the active one when left out.')

/** As in Excel, the name is found whatever its case. */
function sheetOf({ workbook }: SheetToolContext, name: string | undefined): Sheet | string {
  const { sheets, activeSheet } = workbook
  if (name === undefined) return sheets[activeSheet]!
  const wanted = name.toLocaleLowerCase()
  const found = sheets.find((sheet) => sheet.name.toLocaleLowerCase() === wanted)
  return (
    found ?? `There is no sheet named "${name}". The sheets are: ${sheets.map((s) => s.name).join(', ')}.`
  )
}

function usedArea(cells: readonly UsedCell[]): Range | null {
  if (cells.length === 0) return null
  const rows = cells.map((cell) => cell.row)
  const columns = cells.map((cell) => cell.column)
  return {
    fromRow: Math.min(...rows),
    fromColumn: Math.min(...columns),
    toRow: Math.max(...rows),
    toColumn: Math.max(...columns),
  }
}

function outsideProblem(sheet: Sheet, range: Range): string | null {
  if (range.toRow < sheet.rowCount && range.toColumn < sheet.columnCount) return null
  const corner = cellRef(sheet.rowCount - 1, sheet.columnCount - 1)
  return `${describeRange(range)} falls outside the sheet ${sheet.name}, which goes up to ${corner}.`
}

const listSheets = defineTool(
  SheetToolName.ListSheets,
  'List the sheets of the open spreadsheet, where each has cells with content, and the selection.',
  z.object({}),
  (_args, { workbook, selection }) => {
    const lines = workbook.sheets.map((sheet, index) => {
      const area = usedArea(usedCells(sheet))
      const active = index === workbook.activeSheet ? ' (active)' : ''
      return `${index + 1}. ${sheet.name}${active}: ${area === null ? 'empty' : `content in ${describeRange(area)}`}`
    })
    return success([...lines, `Selected on the active sheet: ${describeRange(selection)}`].join('\n'))
  },
)

const readCells = defineTool(
  SheetToolName.ReadCells,
  'Read the cells of a range that have content, one per line: the formula, when there is one, then what the cell shows.',
  z.object({
    range: z
      .string()
      .min(1)
      .max(MAX_REF_LENGTH * 2)
      .describe(RANGE_FORMAT),
    sheet: sheetName,
  }),
  ({ range, sheet: name }, context) => {
    const area = parseRange(range)
    if (area === null) return failure(RANGE_FORMAT)
    const sheet = sheetOf(context, name)
    if (typeof sheet === 'string') return failure(sheet)
    const problem = outsideProblem(sheet, area)
    if (problem !== null) return failure(problem)
    const cells = usedCells(sheet)
      .filter((cell) => rangeContains(area, cell.row, cell.column))
      .sort(byPosition)
    if (cells.length === 0) return success(`${describeRange(area)} on ${sheet.name} is empty.`)
    const lines = [`Sheet: ${sheet.name}`, ...cells.slice(0, MAX_AI_READ_CELLS).map(cellLine)]
    const next = cells[MAX_AI_READ_CELLS]
    if (next !== undefined) {
      lines.push(
        `Only the first ${MAX_AI_READ_CELLS} cells with content are shown; the next is ${cellRef(next.row, next.column)}: read on from its row.`,
      )
    }
    return success(lines.join('\n'))
  },
)

function matches(cell: UsedCell, wanted: string): boolean {
  const texts = [formatCell(cell.cell), cell.cell.formula ?? '']
  return texts.some((text) => text.toLocaleLowerCase().includes(wanted))
}

const findCells = defineTool(
  SheetToolName.FindCells,
  'Find the cells whose formula or shown value contains a text, ignoring case: on one sheet, or on all when none is named.',
  z.object({
    query: z.string().min(1).max(MAX_AI_SEARCH_QUERY_LENGTH),
    sheet: sheetName.describe('Name of the sheet; all of them when left out.'),
    max: z.int().min(1).max(MAX_AI_SEARCH_RESULTS).optional(),
  }),
  ({ query, sheet: name, max = DEFAULT_FIND_RESULTS }, context) => {
    const named = name === undefined ? null : sheetOf(context, name)
    if (typeof named === 'string') return failure(named)
    const wanted = query.toLocaleLowerCase()
    const found: string[] = []
    for (const sheet of named === null ? context.workbook.sheets : [named]) {
      const hits = usedCells(sheet)
        .filter((cell) => matches(cell, wanted))
        .sort(byPosition)
      found.push(...hits.map((cell) => `${quoteSheet(sheet.name)}!${cellLine(cell)}`))
    }
    return success(found.length === 0 ? 'No cell contains that text.' : found.slice(0, max).join('\n'))
  },
)

const cellsToWrite = z
  .array(z.array(z.string().max(MAX_AI_FORMULA_LENGTH)).min(1).max(MAX_AI_WRITE_CELLS))
  .min(1)
  .max(MAX_AI_WRITE_CELLS)
  .refine(
    (rows) => rows.reduce((sum, row) => sum + row.length, 0) <= MAX_AI_WRITE_CELLS,
    `At most ${MAX_AI_WRITE_CELLS} cells per proposal.`,
  )
  .describe('One array per row, one text per cell, as typed: "=SUM(A1:A3)" is a formula, "" clears the cell.')

function formulaProblem(rows: readonly (readonly string[])[], row: number, column: number): string | null {
  for (const [rowOffset, cells] of rows.entries()) {
    for (const [columnOffset, text] of cells.entries()) {
      const problem = text.startsWith('=') ? checkFormula(text) : null
      if (problem !== null) return `${cellRef(row + rowOffset, column + columnOffset)}: ${problem.message}`
    }
  }
  return null
}

const writeCells = defineTool(
  SheetToolName.WriteCells,
  'Propose writing cells from a starting cell. The person sees the proposal over the sheet and applies or discards it; nothing changes until then.',
  z.object({
    start: z.string().min(1).max(MAX_REF_LENGTH).describe('The top left cell, like "B2".'),
    rows: cellsToWrite,
    sheet: sheetName,
    reason: z.string().max(MAX_AI_REASON_LENGTH).optional(),
  }),
  ({ start, rows, sheet: name, reason }, context) => {
    const at = parseRef(start)
    if (at === null) return failure('The start is one cell, like "B2".')
    const sheet = sheetOf(context, name)
    if (typeof sheet === 'string') return failure(sheet)
    const width = Math.max(...rows.map((cells) => cells.length))
    const target = {
      fromRow: at.row,
      fromColumn: at.column,
      toRow: at.row + rows.length - 1,
      toColumn: at.column + width - 1,
    }
    const problem = outsideProblem(sheet, target) ?? formulaProblem(rows, at.row, at.column)
    if (problem !== null) return failure(problem)
    return {
      ...success(
        `The proposal to write ${describeRange(target)} on ${sheet.name} is waiting: the person applies or discards it in Librevia.`,
      ),
      proposal: {
        sheet: sheet.name,
        ...at,
        rows,
        ...(reason === undefined || reason === '' ? {} : { reason }),
      },
    }
  },
)

const TOOLS = toolbox<SheetToolContext, SheetProposal>([listSheets, readCells, findCells, writeCells])

export const SHEET_TOOLS: readonly AiToolDefinition[] = TOOLS.definitions

export const isSheetTool = (name: string): boolean => TOOLS.has(name)

/** Reads only the workbook it is given; a write is only a proposal, which the person applies. */
export function runSheetTool(call: AiToolCall, context: SheetToolContext): SheetOutcome {
  return TOOLS.run(call, context)
}
