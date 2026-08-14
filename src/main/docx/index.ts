/**
 * DOCX no processo main: conversa com o sidecar e guarda os bytes originais.
 *
 * Os bytes ficam **aqui**, e não no sidecar, por decisão de desenho
 * (docs/02-docx-cirurgico.md): o sidecar é sem estado, então a morte dele não
 * custa a capacidade de gravar cirurgicamente. Guardá-los lá quebraria a
 * promessa da Fase 3.5.
 */

import { readFile } from 'node:fs/promises'
import { z } from 'zod'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import type { LossInventory } from '@shared/types.js'
import type { SidecarClient } from '../sidecar/client.js'
import { SidecarMethod } from '../sidecar/protocol.js'

const inventorySchema = z.object({
  invisible: z.array(z.string()).default([]),
  lost: z.array(z.string()).default([]),
})

const openResultSchema = z.object({
  model: z.object({ page: z.unknown(), doc: z.unknown() }),
  inventory: inventorySchema,
})

const saveResultSchema = z.object({
  inventory: inventorySchema,
  preservedBlocks: z.number().int().nonnegative(),
  rewrittenBlocks: z.number().int().nonnegative(),
})

/**
 * Os bytes do documento aberto, como estavam no disco na hora de abrir.
 *
 * Uma entrada só: o aplicativo edita um documento por vez. E é de propósito que
 * a fonte de verdade seja *o arquivo como você o abriu*, e não como ele está no
 * disco agora — que pode ter mudado por outra mão no meio da edição.
 */
let openedOriginal: { path: string; bytes: Buffer } | null = null

export function forgetOpenedDocx(): void {
  openedOriginal = null
}

export interface OpenedDocx {
  /** O modelo já no envelope `.sdoc`, para o renderer seguir por um caminho só. */
  readonly content: string
  readonly inventory: LossInventory
}

export async function openDocx(client: SidecarClient, path: string): Promise<OpenedDocx> {
  let bytes: Buffer
  try {
    bytes = await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura')
  }

  const reply = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
  const parsed = openResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      'Não foi possível ler este documento do Word. O arquivo pode estar danificado.',
      'docx.open fora do contrato',
    )
  }

  openedOriginal = { path, bytes }

  return {
    content: JSON.stringify({ format: 'sdoc', version: 1, ...parsed.data.model }),
    inventory: parsed.data.inventory,
  }
}

export interface SavedDocx {
  readonly bytes: Uint8Array
  readonly inventory: LossInventory
}

/**
 * Grava o modelo por cima do pacote original.
 *
 * `sdocContent` é o mesmo JSON que o renderer manda para qualquer destino — o
 * envelope é descascado aqui para que o renderer não precise saber que DOCX
 * existe.
 */
export async function saveDocx(client: SidecarClient, sdocContent: string): Promise<SavedDocx> {
  const model = unwrapSdoc(sdocContent)

  const original = openedOriginal?.bytes
  if (original === undefined) {
    throw new AppError(
      ErrorCode.UnsupportedFormat,
      'Só é possível salvar em .docx um documento que foi aberto a partir de um arquivo .docx. Salve como .sdoc ou abra um documento do Word primeiro.',
    )
  }

  const reply = await client.request(SidecarMethod.DocxSave, model, new Uint8Array(original))
  const parsed = saveResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      'Não foi possível gravar o documento do Word. O arquivo original não foi alterado.',
      'docx.save fora do contrato',
    )
  }

  console.info(`[docx] preservados ${parsed.data.preservedBlocks}, reescritos ${parsed.data.rewrittenBlocks}`)

  return { bytes: reply.binary, inventory: parsed.data.inventory }
}

function unwrapSdoc(content: string): unknown {
  let parsed: unknown
  try {
    parsed = JSON.parse(content)
  } catch {
    throw new AppError(ErrorCode.Internal, 'O documento em edição está em estado inconsistente.')
  }

  const envelope = z.object({ page: z.unknown(), doc: z.unknown() }).safeParse(parsed)
  if (!envelope.success) {
    throw new AppError(ErrorCode.Internal, 'O documento em edição está em estado inconsistente.')
  }

  return { page: envelope.data.page, doc: envelope.data.doc }
}
