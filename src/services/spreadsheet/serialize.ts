import { z } from 'zod'
import { AppError, ErrorCode } from '@shared/errors.js'
import { DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT, createEmptyWorkbook, type WorkbookModel } from './model.js'

/**
 * Formato interno `.ssheet`.
 *
 * Mesma ideia do `.sdoc`: o modelo gravado como está, num JSON. Serve para
 * salvar e reabrir **sem perda nenhuma** enquanto o XLSX não chega (Fase 7).
 *
 * O mapa de células é esparso também no arquivo — uma planilha de 10 mil linhas
 * com trinta valores gera trinta entradas, não dez mil linhas vazias.
 */
export const SSHEET_FORMAT = 'ssheet'
export const SSHEET_VERSION = 1

const cellStyleSchema = z.object({
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  color: z.string().max(32).optional(),
  background: z.string().max(32).optional(),
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

/**
 * Dimensões por índice.
 *
 * As chaves saem do JSON como texto e voltam como número — `z.coerce` faz a
 * conversão, senão `columnWidths[3]` nunca encontraria a entrada `"3"`.
 */
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

/**
 * Lê um `.ssheet`.
 *
 * Arquivo corrompido ou de versão futura produz uma frase que o usuário
 * entende, e não um erro de JSON.
 */
export function parseWorkbook(text: string): WorkbookModel {
  let raw: unknown
  try {
    raw = JSON.parse(text)
  } catch {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Esta planilha não pôde ser lida: o conteúdo está corrompido ou não é uma planilha válida.',
    )
  }

  const parsed = ssheetSchema.safeParse(raw)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Este arquivo não é uma planilha válida deste aplicativo.',
    )
  }

  if (parsed.data.version > SSHEET_VERSION) {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Esta planilha foi criada por uma versão mais recente do aplicativo. Atualize para abri-la.',
    )
  }

  // Aba ativa fora do intervalo não impede a leitura: os dados valem mais que
  // a lembrança de qual aba estava aberta.
  const activeSheet = parsed.data.activeSheet < parsed.data.sheets.length ? parsed.data.activeSheet : 0

  return { sheets: parsed.data.sheets, activeSheet }
}

export function isSpreadsheetFile(text: string): boolean {
  return text.trimStart().startsWith('{') && text.includes(`"${SSHEET_FORMAT}"`)
}

export function createEmptySpreadsheetFile(): string {
  return serializeWorkbook(createEmptyWorkbook())
}
