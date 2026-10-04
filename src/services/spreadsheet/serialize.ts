import { z } from 'zod'
import { AppError, ErrorCode } from '@shared/errors.js'
import { translate, Language } from '@shared/i18n/index.js'
import { DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT, createEmptyWorkbook, type WorkbookModel } from './model.js'
import { MAX_COLOR_LENGTH } from '@shared/limits.js'

export const SSHEET_FORMAT = 'ssheet'
export const SSHEET_VERSION = 1

const cellStyleSchema = z.object({
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  color: z.string().max(MAX_COLOR_LENGTH).optional(),
  background: z.string().max(MAX_COLOR_LENGTH).optional(),
  align: z.enum(['left', 'center', 'right']).optional(),
  format: z.enum(['general', 'text', 'number', 'currency', 'percent', 'date']).optional(),
  decimals: z.number().int().min(0).max(10).optional(),
  borders: z
    .array(z.enum(['top', 'right', 'bottom', 'left']))
    .max(4)
    .optional(),
})

const cellSchema = z.object({
  value: z.union([z.string(), z.number(), z.boolean()]).optional(),
  formula: z.string().max(8000).optional(),
  style: cellStyleSchema.optional(),
})

/** As chaves saem do JSON como texto: `z.coerce` as devolve a número. */
const dimensionsSchema = z.record(z.coerce.number().int().nonnegative(), z.number().positive().max(4000))

const sheetSchema = z.object({
  name: z.string().min(1).max(120),
  cells: z.record(z.string().max(16), cellSchema),
  columnWidths: dimensionsSchema.default({}),
  rowHeights: dimensionsSchema.default({}),
  frozenRows: z.number().int().min(0).max(100).default(0),
  frozenColumns: z.number().int().min(0).max(100).default(0),
  rowCount: z.number().int().positive().max(1_000_000).default(DEFAULT_ROW_COUNT),
  columnCount: z.number().int().positive().max(16_384).default(DEFAULT_COLUMN_COUNT),
})

const ssheetSchema = z.object({
  format: z.literal(SSHEET_FORMAT),
  version: z.number().int().positive(),
  sheets: z.array(sheetSchema).min(1).max(200),
  activeSheet: z.number().int().nonnegative().default(0),
})

export function serializeWorkbook(model: WorkbookModel): string {
  return JSON.stringify(
    { format: SSHEET_FORMAT, version: SSHEET_VERSION, sheets: model.sheets, activeSheet: model.activeSheet },
    null,
    2,
  )
}

/** Arquivo corrompido ou de versão futura produz uma frase, e não um erro de JSON. */
export function parseWorkbook(text: string, language: Language = Language.Portuguese): WorkbookModel {
  let raw: unknown
  try {
    raw = JSON.parse(text)
  } catch {
    throw new AppError(ErrorCode.UnsupportedFormat, translate(language, 'spreadsheet.error.corrupt'))
  }

  const parsed = ssheetSchema.safeParse(raw)
  if (!parsed.success) {
    throw new AppError(ErrorCode.UnsupportedFormat, translate(language, 'spreadsheet.error.invalid'))
  }

  if (parsed.data.version > SSHEET_VERSION) {
    throw new AppError(ErrorCode.UnsupportedFormat, translate(language, 'spreadsheet.error.newerVersion'))
  }

  // Aba ativa fora do intervalo não impede a leitura.
  const activeSheet = parsed.data.activeSheet < parsed.data.sheets.length ? parsed.data.activeSheet : 0

  return { sheets: parsed.data.sheets, activeSheet }
}

export function isSpreadsheetFile(text: string): boolean {
  return text.trimStart().startsWith('{') && text.includes(`"${SSHEET_FORMAT}"`)
}

export function createEmptySpreadsheetFile(): string {
  return serializeWorkbook(createEmptyWorkbook())
}
