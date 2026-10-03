/**
 * Como o DOCX, os bytes originais moram aqui. A fórmula é traduzida **aqui**
 * (`SUM(A1,B1)` ↔ `SOMA(A1;B1)`), e não no sidecar: o analisador já existe do
 * lado TypeScript.
 */

import { readFile } from 'node:fs/promises'
import { z } from 'zod'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import type { LossInventory } from '@shared/types.js'
import { isKnownFunction } from '@services/spreadsheet/formula/functions/index.js'
import { fromXlsxFormula, toXlsxFormula } from '@services/spreadsheet/formula/interop.js'
import { TokenKind, tokenize } from '@services/spreadsheet/formula/tokenize.js'
import type { CellMap, Sheet, WorkbookModel } from '@services/spreadsheet/model.js'
import {
  SSHEET_FORMAT,
  SSHEET_VERSION,
  parseWorkbook,
  serializeWorkbook,
} from '@services/spreadsheet/serialize.js'
import type { SidecarClient } from '../sidecar/client.js'
import { SidecarMethod } from '../sidecar/protocol.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

const inventorySchema = z.object({
  invisible: z.array(z.string()).default([]),
  lost: z.array(z.string()).default([]),
  structural: z.array(z.string()).default([]),
})

const openResultSchema = z.object({
  workbook: z.unknown(),
  inventory: inventorySchema,
})

const saveResultSchema = z.object({
  sheets: z.number().int().nonnegative(),
  cellsWritten: z.number().int().nonnegative(),
  cellsCleared: z.number().int().nonnegative(),
  cellsPreserved: z.number().int().nonnegative(),
})

let openedOriginal: { path: string; bytes: Buffer } | null = null

export function forgetOpenedXlsx(): void {
  openedOriginal = null
}

/** Ver `adoptDocxOriginal`. */
export async function adoptXlsxOriginal(path: string): Promise<boolean> {
  try {
    openedOriginal = { path, bytes: await readFile(path) }
    return true
  } catch {
    openedOriginal = null
    return false
  }
}

export interface OpenedXlsx {
  readonly content: string
  readonly inventory: LossInventory
}

export async function openXlsx(client: SidecarClient, path: string): Promise<OpenedXlsx> {
  let bytes: Buffer
  try {
    bytes = await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura', editorPreferences().language)
  }

  // Os dois tempos dizem se a demora foi do sidecar ou da conversão.
  const startedAt = Date.now()
  const reply = await client.request(SidecarMethod.XlsxOpen, {}, new Uint8Array(bytes))
  const readAt = Date.now()

  const parsed = openResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.xlsx.cannotRead'), t('errors.xlsx.openContract'))
  }

  const model = translate(toModel(parsed.data.workbook), fromXlsxFormula)
  openedOriginal = { path, bytes }

  const cells = model.sheets.reduce((total, sheet) => total + Object.keys(sheet.cells).length, 0)
  console.info(
    `[xlsx] abertas ${cells} celulas — servico ${readAt - startedAt} ms, conversao ${Date.now() - readAt} ms`,
  )

  return {
    content: serializeWorkbook(model),
    inventory: withUncalculated(parsed.data.inventory, model),
  }
}

export interface SavedXlsx {
  readonly bytes: Uint8Array
  readonly inventory: LossInventory
}

/** Sem original, o sidecar monta um pacote novo: uma planilha é grade, valor e fórmula, sem nada a perder. */
export async function saveXlsx(client: SidecarClient, ssheetContent: string): Promise<SavedXlsx> {
  const model = translate(readSsheet(ssheetContent), toXlsxFormula)
  const original = openedOriginal?.bytes

  const reply = await client.request(
    SidecarMethod.XlsxSave,
    { sheets: model.sheets, activeSheet: model.activeSheet },
    original === undefined ? new Uint8Array(0) : new Uint8Array(original),
  )

  const parsed = saveResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.xlsx.cannotSave'), t('errors.xlsx.saveContract'))
  }

  console.info(
    `[xlsx] escritas ${parsed.data.cellsWritten}, limpas ${parsed.data.cellsCleared}, preservadas ${parsed.data.cellsPreserved}`,
  )

  return { bytes: reply.binary, inventory: { invisible: [], lost: [], structural: [] } }
}

function toModel(workbook: unknown): WorkbookModel {
  const envelope = { format: SSHEET_FORMAT, version: SSHEET_VERSION, ...(workbook as object) }
  try {
    return parseWorkbook(JSON.stringify(envelope))
  } catch {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.xlsx.cannotRead'), t('errors.xlsx.invalidSchema'))
  }
}

function readSsheet(content: string): WorkbookModel {
  try {
    return parseWorkbook(content)
  } catch {
    throw new AppError(ErrorCode.UnsupportedFormat, t('errors.xlsx.onlySpreadsheets'))
  }
}

function translate(model: WorkbookModel, convert: (formula: string) => string): WorkbookModel {
  return { ...model, sheets: model.sheets.map((sheet) => translateSheet(sheet, convert)) }
}

function translateSheet(sheet: Sheet, convert: (formula: string) => string): Sheet {
  const cells: CellMap = {}
  for (const [reference, cell] of Object.entries(sheet.cells)) {
    cells[reference] = cell.formula === undefined ? cell : { ...cell, formula: convert(cell.formula) }
  }
  return { ...sheet, cells }
}

/** Invisibilidade, e não perda: a fórmula volta intacta, mas a célula mostra `#NOME?`. */
function withUncalculated(inventory: LossInventory, model: WorkbookModel): LossInventory {
  const unknown = new Set<string>()

  for (const sheet of model.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      if (cell.formula === undefined) continue
      for (const name of functionsIn(cell.formula)) {
        if (!isKnownFunction(name)) unknown.add(name)
      }
    }
  }

  if (unknown.size === 0) return inventory

  const names = [...unknown].sort((a, b) => a.localeCompare(b, 'pt-BR')).join(', ')
  return {
    ...inventory,
    invisible: [...inventory.invisible, t('errors.xlsx.unknownFunctions', { names })],
  }
}

function functionsIn(formula: string): string[] {
  try {
    return tokenize(formula.startsWith('=') ? formula.slice(1) : formula)
      .filter((token) => token.kind === TokenKind.Name)
      .map((token) => token.text)
  } catch {
    // A célula já vai mostrar o erro.
    return []
  }
}
