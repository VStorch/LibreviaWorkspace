/**
 * **Bytes entram, JSON sai**, num quadro binário que as duas linguagens montam
 * sem framework:
 *
 * ```text
 *   offset 0   uint32 BE   bytes de JSON
 *   offset 4   uint32 BE   bytes de binário
 *   offset 8   ...         JSON em UTF-8
 *   depois     ...         binário cru
 * ```
 *
 * O binário viaja fora do JSON: Base64 custaria 33% a mais em documentos de até
 * 20 MB.
 */

import { z } from 'zod'
import { AppError, ErrorCode } from '@shared/errors.js'
import { t } from '../i18n.js'

export const FRAME_HEADER_BYTES = 8

/**
 * Contra um sidecar corrompido anunciando um quadro absurdo. O teto do JSON é
 * generoso porque as imagens do DOCX vão como data URI.
 */
export const MAX_JSON_BYTES = 64 * 1024 * 1024
export const MAX_BINARY_BYTES = 64 * 1024 * 1024

export interface Frame {
  readonly json: unknown
  readonly binary: Uint8Array
}

export const EMPTY_BINARY = new Uint8Array(0)

export function encodeFrame(json: unknown, binary: Uint8Array = EMPTY_BINARY): Uint8Array {
  const jsonBytes = new TextEncoder().encode(JSON.stringify(json))

  const frame = new Uint8Array(FRAME_HEADER_BYTES + jsonBytes.length + binary.length)
  const header = new DataView(frame.buffer, 0, FRAME_HEADER_BYTES)
  header.setUint32(0, jsonBytes.length, false)
  header.setUint32(4, binary.length, false)

  frame.set(jsonBytes, FRAME_HEADER_BYTES)
  frame.set(binary, FRAME_HEADER_BYTES + jsonBytes.length)
  return frame
}

/** Um pipe não preserva fronteiras: um quadro chega partido, e vários chegam juntos. */
export class FrameReader {
  // Anotado: `new Uint8Array(0)` inferiria um tipo mais estreito que os pedaços do pipe.
  #buffer: Uint8Array = EMPTY_BINARY

  push(chunk: Uint8Array): Frame[] {
    this.#buffer = concat(this.#buffer, chunk)

    const frames: Frame[] = []
    for (;;) {
      const frame = this.#take()
      if (frame === undefined) return frames
      frames.push(frame)
    }
  }

  #take(): Frame | undefined {
    if (this.#buffer.length < FRAME_HEADER_BYTES) return undefined

    const header = new DataView(this.#buffer.buffer, this.#buffer.byteOffset, FRAME_HEADER_BYTES)
    const jsonLength = header.getUint32(0, false)
    const binaryLength = header.getUint32(4, false)

    if (jsonLength > MAX_JSON_BYTES || binaryLength > MAX_BINARY_BYTES) {
      throw new AppError(
        ErrorCode.SidecarFailed,
        t('errors.sidecar.unexpectedResponse'),
        t('errors.sidecar.frameAnnounce', { jsonLength, binaryLength }),
      )
    }

    const total = FRAME_HEADER_BYTES + jsonLength + binaryLength
    if (this.#buffer.length < total) return undefined

    const jsonBytes = this.#buffer.subarray(FRAME_HEADER_BYTES, FRAME_HEADER_BYTES + jsonLength)
    // Cópia: o binário sobrevive ao buffer, que será fatiado.
    const binary = this.#buffer.slice(FRAME_HEADER_BYTES + jsonLength, total)
    this.#buffer = this.#buffer.slice(total)

    return { json: parseJson(jsonBytes), binary }
  }
}

function parseJson(bytes: Uint8Array): unknown {
  try {
    return JSON.parse(new TextDecoder().decode(bytes))
  } catch {
    throw new AppError(
      ErrorCode.SidecarFailed,
      t('errors.sidecar.unexpectedResponse'),
      t('errors.sidecar.invalidJson'),
    )
  }
}

function concat(left: Uint8Array, right: Uint8Array): Uint8Array {
  if (left.length === 0) return right
  if (right.length === 0) return left

  const merged = new Uint8Array(left.length + right.length)
  merged.set(left, 0)
  merged.set(right, left.length)
  return merged
}


/**
 * `health` prova que o processo sobe; `diagnostics.echo`, que o binário atravessa.
 * `docx.save` recebe os bytes originais, e não uma sessão, porque o sidecar é sem
 * estado. `docx.create` devolve o pacote mínimo do documento que nasceu no editor.
 */
export const SidecarMethod = {
  Health: 'health',
  Echo: 'diagnostics.echo',
  DocxOpen: 'docx.open',
  DocxSave: 'docx.save',
  DocxCreate: 'docx.create',
  XlsxOpen: 'xlsx.open',
  XlsxSave: 'xlsx.save',
} as const

export type SidecarMethod = (typeof SidecarMethod)[keyof typeof SidecarMethod]

export interface SidecarRequest {
  readonly id: number
  readonly method: SidecarMethod
  readonly params: unknown
}

/** O sidecar fala português com o usuário; o código do erro é dele. */
const sidecarErrorSchema = z.object({
  code: z.string().min(1).max(64),
  message: z.string().min(1).max(500),
  detail: z.string().max(500).optional(),
})

const responseSchema = z.discriminatedUnion('ok', [
  z.object({
    id: z.number().int().nonnegative(),
    ok: z.literal(true),
    // O .NET omite propriedade nula, e operação só com binário é o caso normal.
    result: z.unknown().optional(),
  }),
  z.object({ id: z.number().int().nonnegative(), ok: z.literal(false), error: sidecarErrorSchema }),
])

export type SidecarResponse = z.infer<typeof responseSchema>

/** O sidecar é outro processo: pode estar antigo, trocado ou corrompido. */
export function parseResponse(json: unknown): SidecarResponse {
  const parsed = responseSchema.safeParse(json)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      t('errors.sidecar.unexpectedResponse'),
      t('errors.sidecar.contractViolation'),
    )
  }
  return parsed.data
}

export const healthResultSchema = z.object({
  name: z.string(),
  version: z.string(),
  runtime: z.string(),
})

export type HealthResult = z.infer<typeof healthResultSchema>
