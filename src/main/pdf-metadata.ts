/**
 * A PDF's properties go through the sidecar, which parses the file away from main. Saving appends
 * the sidecar's incremental update to the bytes on disk: everything before it stays as it was.
 */

import { readFile, stat } from 'node:fs/promises'
import { z } from 'zod'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { pdfInspectionSchema } from '@shared/schemas.js'
import type { DocumentProperties } from '@services/file/properties.js'
import { writeFileAtomic } from './fs/atomic-write.js'
import { t } from './i18n.js'
import { editorPreferences } from './preferences.js'
import type { SidecarClient } from './sidecar/client.js'
import { MAX_BINARY_BYTES, SidecarMethod } from './sidecar/protocol.js'

export type PdfInspection = z.infer<typeof pdfInspectionSchema>

const setMetadataResultSchema = z.object({ appended: z.number().int().nonnegative() })

export async function inspectPdf(client: SidecarClient, path: string): Promise<PdfInspection> {
  const reply = await client.request(SidecarMethod.PdfInspect, {}, await readPdf(path))
  const parsed = pdfInspectionSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.pdf.cannotRead'), t('errors.pdf.inspectContract'))
  }
  return parsed.data
}

/** Nothing is written when no field differs from the file's. */
export async function savePdfProperties(
  client: SidecarClient,
  path: string,
  properties: DocumentProperties,
): Promise<void> {
  const original = await readPdf(path)
  const reply = await client.request(SidecarMethod.PdfSetMetadata, properties, original)
  if (!setMetadataResultSchema.safeParse(reply.result).success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.pdf.cannotSave'), t('errors.pdf.saveContract'))
  }
  if (reply.binary.length === 0) return

  const updated = new Uint8Array(original.length + reply.binary.length)
  updated.set(original, 0)
  updated.set(reply.binary, original.length)
  await writeFileAtomic(path, updated)
}

/** The whole file crosses to the sidecar, which reads objects anywhere in it. */
async function readPdf(path: string): Promise<Uint8Array> {
  try {
    if ((await stat(path)).size > MAX_BINARY_BYTES) {
      const limit = Math.round(MAX_BINARY_BYTES / (1024 * 1024))
      throw new AppError(ErrorCode.FileTooLarge, t('errors.pdf.tooLarge', { limit }))
    }
    return new Uint8Array(await readFile(path))
  } catch (cause) {
    if (cause instanceof AppError) throw cause
    throw fromFileSystemError(cause, 'leitura', editorPreferences().language)
  }
}
