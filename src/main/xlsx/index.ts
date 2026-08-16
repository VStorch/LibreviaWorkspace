/**
 * XLSX no processo main: conversa com o sidecar e guarda os bytes originais.
 *
 * Espelha `main/docx/index.ts`, e pelo mesmo motivo: o sidecar é sem estado, e
 * os bytes originais moram aqui para que a morte dele não custe a capacidade de
 * gravar cirurgicamente.
 *
 * A diferença fica na fórmula. O arquivo guarda `SUM(A1,B1)`; o aplicativo usa
 * `SOMA(A1;B1)`. A tradução acontece **aqui**, na fronteira, e não no sidecar:
 * o analisador já existe do lado TypeScript, e uma segunda gramática em C#
 * seria duas coisas para manter em acordo.
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

const inventorySchema = z.object({
  invisible: z.array(z.string()).default([]),
  lost: z.array(z.string()).default([]),
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

/** Os bytes da planilha aberta, como estavam no disco na hora de abrir. */
let openedOriginal: { path: string; bytes: Buffer } | null = null

export function forgetOpenedXlsx(): void {
  openedOriginal = null
}

/** Reata o vínculo com o pacote original depois de uma recuperação. Ver o DOCX. */
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
  /** O modelo já no envelope `.ssheet`, para o renderer seguir por um caminho só. */
  readonly content: string
  readonly inventory: LossInventory
}

export async function openXlsx(client: SidecarClient, path: string): Promise<OpenedXlsx> {
  let bytes: Buffer
  try {
    bytes = await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura')
  }

  // Cronometrado porque "demorou para abrir" é a reclamação mais difícil de
  // diagnosticar depois: sem os dois números, não dá para saber se o tempo foi
  // do serviço de formatos ou da conversão deste lado.
  const startedAt = Date.now()
  const reply = await client.request(SidecarMethod.XlsxOpen, {}, new Uint8Array(bytes))
  const readAt = Date.now()

  const parsed = openResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      'Não foi possível ler esta planilha do Excel. O arquivo pode estar danificado.',
      'xlsx.open fora do contrato',
    )
  }

  const model = translate(toModel(parsed.data.workbook), fromXlsxFormula)
  openedOriginal = { path, bytes }

  const cells = model.sheets.reduce((total, sheet) => total + Object.keys(sheet.cells).length, 0)
  console.info(
    `[xlsx] abertas ${cells} células — serviço ${readAt - startedAt} ms, conversão ${Date.now() - readAt} ms`,
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

/**
 * Grava o modelo por cima do pacote original.
 *
 * Sem original — planilha criada aqui e salva como `.xlsx` pela primeira vez —
 * o sidecar monta um pacote novo. É a diferença para o DOCX, que se recusa a
 * criar do zero: uma planilha é grade, valor e fórmula, e um arquivo montado
 * assim não perde nada de um original que não existe.
 */
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
    throw new AppError(
      ErrorCode.SidecarFailed,
      'Não foi possível gravar a planilha do Excel. O arquivo original não foi alterado.',
      'xlsx.save fora do contrato',
    )
  }

  console.info(
    `[xlsx] escritas ${parsed.data.cellsWritten}, limpas ${parsed.data.cellsCleared}, preservadas ${parsed.data.cellsPreserved}`,
  )

  return { bytes: reply.binary, inventory: { invisible: [], lost: [] } }
}

/** O modelo que veio do sidecar, conferido pelo mesmo esquema do `.ssheet`. */
function toModel(workbook: unknown): WorkbookModel {
  const envelope = { format: SSHEET_FORMAT, version: SSHEET_VERSION, ...(workbook as object) }
  try {
    return parseWorkbook(JSON.stringify(envelope))
  } catch {
    throw new AppError(
      ErrorCode.SidecarFailed,
      'Não foi possível ler esta planilha do Excel. O arquivo pode estar danificado.',
      'xlsx.open devolveu um modelo fora do esquema',
    )
  }
}

function readSsheet(content: string): WorkbookModel {
  try {
    return parseWorkbook(content)
  } catch {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Só é possível salvar em .xlsx uma planilha. Para um documento de texto, use .docx ou .sdoc.',
    )
  }
}

/** Reescreve toda fórmula do modelo, deixando o resto como está. */
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

/**
 * Acrescenta ao inventário as funções que o motor não calcula.
 *
 * É invisibilidade, não perda: a fórmula continua no arquivo e volta intacta ao
 * gravar. O que muda é a tela, onde a célula mostra `#NOME?` em vez do valor que
 * o Excel havia calculado — e sem aviso o usuário concluiria que a planilha
 * quebrou.
 */
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
    invisible: [...inventory.invisible, `funções que este aplicativo não calcula: ${names}`],
  }
}

function functionsIn(formula: string): string[] {
  try {
    return tokenize(formula.startsWith('=') ? formula.slice(1) : formula)
      .filter((token) => token.kind === TokenKind.Name)
      .map((token) => token.text)
  } catch {
    // Fórmula que nem se consegue separar em símbolos já vai mostrar erro na
    // célula. Um segundo aviso sobre ela não ajudaria em nada.
    return []
  }
}
