/**
 * Contrato entre o processo main e o sidecar .NET.
 *
 * A costura é deliberadamente burra: **bytes entram, JSON sai**. Nenhum objeto
 * de biblioteca, nenhuma classe, nenhuma referência compartilhada — só um
 * quadro binário que as duas linguagens sabem montar sem framework nenhum.
 * É isso que permite trocar a implementação do outro lado sem tocar no app.
 *
 * ## Formato do quadro
 *
 * ```text
 *   offset 0   uint32 BE   bytes de JSON
 *   offset 4   uint32 BE   bytes de binário
 *   offset 8   ...         JSON em UTF-8
 *   depois     ...         binário cru
 * ```
 *
 * O binário viaja **fora** do JSON de propósito. Base64 custaria 33% a mais em
 * cima de documentos de até 20 MB, e a conversão apareceria no tempo de abrir
 * cada arquivo.
 */

import { z } from 'zod'
import { AppError, ErrorCode } from '@shared/errors.js'

export const FRAME_HEADER_BYTES = 8

/**
 * Tetos de sanidade. Não são política de produto: existem para que um sidecar
 * corrompido — ou trocado por outra coisa — não consiga nos fazer alocar
 * gigabytes anunciando um quadro absurdo.
 *
 * O teto do JSON é generoso porque o modelo de um DOCX carrega as imagens como
 * data URI: um arquivo de 6,6 MB cheio de capturas de tela vira ~9 MB de JSON,
 * e o limite de 20 MB por arquivo (`MAX_FILE_BYTES`) chega perto de 30 MB
 * depois do base64.
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

/**
 * Remonta quadros a partir de pedaços arbitrários de stdout.
 *
 * Um pipe não preserva fronteiras de mensagem: um quadro pode chegar partido em
 * cinco pedaços, e cinco quadros podem chegar num pedaço só. Tratar cada chunk
 * como se fosse uma mensagem é o bug clássico desse tipo de integração, e ele
 * só aparece sob carga — exatamente quando o documento é grande.
 */
export class FrameReader {
  // Anotado: `new Uint8Array(0)` infere `Uint8Array<ArrayBuffer>`, mais estreito
  // que os pedaços que chegam do pipe.
  #buffer: Uint8Array = EMPTY_BINARY

  /** Consome um pedaço e devolve os quadros que ficaram completos com ele. */
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
        'O serviço de formatos respondeu de forma inesperada. A operação não foi concluída.',
        `quadro anuncia ${jsonLength} bytes de JSON e ${binaryLength} de binário`,
      )
    }

    const total = FRAME_HEADER_BYTES + jsonLength + binaryLength
    if (this.#buffer.length < total) return undefined

    const jsonBytes = this.#buffer.subarray(FRAME_HEADER_BYTES, FRAME_HEADER_BYTES + jsonLength)
    // Cópia proposital: o binário sobrevive ao buffer, que será fatiado a seguir.
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
      'O serviço de formatos respondeu de forma inesperada. A operação não foi concluída.',
      'JSON inválido no quadro',
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

// --- Mensagens -------------------------------------------------------------

/**
 * Métodos disponíveis.
 *
 * `health` prova que o processo sobe; `diagnostics.echo` prova que o binário
 * atravessa inteiro. Os dois de DOCX são sem estado: `docx.save` recebe os
 * bytes originais no lugar de um identificador de sessão, para que a morte do
 * sidecar não custe a gravação cirúrgica.
 */
export const SidecarMethod = {
  Health: 'health',
  Echo: 'diagnostics.echo',
  DocxOpen: 'docx.open',
  DocxSave: 'docx.save',
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
    // Opcional de propósito: o serializador do .NET omite propriedade nula, e
    // operação sem valor de retorno — só binário — é o caso normal, não o
    // excepcional. Ausente e nulo significam a mesma coisa aqui.
    result: z.unknown().optional(),
  }),
  z.object({ id: z.number().int().nonnegative(), ok: z.literal(false), error: sidecarErrorSchema }),
])

export type SidecarResponse = z.infer<typeof responseSchema>

/**
 * Valida a resposta antes de qualquer uso.
 *
 * O sidecar é código nosso, mas é **outro processo**: pode estar numa versão
 * antiga, pode ter sido substituído no disco, pode estar corrompido. Confiar na
 * forma do que ele devolve seria o mesmo erro de confiar no renderer.
 */
export function parseResponse(json: unknown): SidecarResponse {
  const parsed = responseSchema.safeParse(json)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      'O serviço de formatos respondeu de forma inesperada. A operação não foi concluída.',
      'resposta fora do contrato',
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
