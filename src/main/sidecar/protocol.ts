/**
 * **Bytes in, JSON out**, in a binary frame both languages build without a framework:
 *
 * ```text
 *   offset 0   uint32 BE   JSON byte count
 *   offset 4   uint32 BE   binary byte count
 *   offset 8   ...         UTF-8 JSON
 *   then       ...         raw binary
 * ```
 *
 * The binary travels outside the JSON: Base64 would cost 33% more on documents of up to 20 MB.
 */

import { z } from 'zod'
import { AppError, ErrorCode } from '@shared/errors.js'
import { t } from '../i18n.js'

export const FRAME_HEADER_BYTES = 8

/**
 * Against a corrupt sidecar announcing an absurd frame. The JSON cap is generous because DOCX
 * images travel as data URIs.
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

/** A pipe does not preserve boundaries: a frame arrives split, and several arrive together. */
export class FrameReader {
  // Annotated: `new Uint8Array(0)` would infer a narrower type than the pipe chunks.
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
    // A copy: the binary outlives the buffer, which will be sliced.
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
 * `health` proves the process starts; `diagnostics.echo`, that the binary gets through. `docx.save`
 * receives the original bytes, not a session, because the sidecar is stateless. `docx.create`
 * returns the minimal package for a document born in the editor.
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

/** The sidecar speaks Portuguese to the user; the error code is its own. */
const sidecarErrorSchema = z.object({
  code: z.string().min(1).max(64),
  message: z.string().min(1).max(500),
  detail: z.string().max(500).optional(),
})

const responseSchema = z.discriminatedUnion('ok', [
  z.object({
    id: z.number().int().nonnegative(),
    ok: z.literal(true),
    // .NET omits null properties, and a binary-only operation is the normal case.
    result: z.unknown().optional(),
  }),
  z.object({ id: z.number().int().nonnegative(), ok: z.literal(false), error: sidecarErrorSchema }),
])

export type SidecarResponse = z.infer<typeof responseSchema>

/** The sidecar is another process: it may be old, swapped or corrupt. */
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
