/**
 * An ODT opens as a file of its own: main keeps the original bytes, as for a `.docx`, and saving
 * goes back into that package through the sidecar, which rewrites only the edited blocks.
 */

import { withoutUnsafeLinks } from '@services/document/link.js'
import { AppError, ErrorCode } from '@shared/errors.js'
import {
  openResultSchema,
  readSdocModel,
  saveParamsOf,
  saveResultSchema,
  sdocContentOf,
  type OpenedDocx,
  type SavedDocx,
} from '../docx/index.js'
import { readFileBytes } from '../fs/read-bytes.js'
import { OriginalFormat, originalOf, rememberOriginal } from '../opened-original.js'
import type { SidecarClient } from '../sidecar/client.js'
import { SidecarMethod } from '../sidecar/protocol.js'
import { t } from '../i18n.js'

export async function openOdt(client: SidecarClient, path: string): Promise<OpenedDocx> {
  return decodeOdt(client, path, await readFileBytes(path))
}

export async function decodeOdt(client: SidecarClient, path: string, bytes: Buffer): Promise<OpenedDocx> {
  const reply = await client.request(SidecarMethod.OdtOpen, {}, new Uint8Array(bytes))
  const parsed = openResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.odt.cannotRead'), t('errors.odt.openContract'))
  }

  rememberOriginal(path, bytes, OriginalFormat.Odt)
  const { model, inventory } = parsed.data
  return { content: sdocContentOf({ ...model, doc: withoutUnsafeLinks(model.doc) }), inventory }
}

/** Whether `.odt` is a destination for the document that came from `origin`. */
export function canSaveAsOdt(origin: string | null): boolean {
  return originalOf(origin, OriginalFormat.Odt) !== null
}

/**
 * Only over the `.odt` the document came from: a document without one is exported, not saved
 * (`export-odt.ts`), until the sidecar can create the package.
 */
export function assertOdtDestination(origin: string | null): Buffer {
  const original = originalOf(origin, OriginalFormat.Odt)
  if (original === null) throw new AppError(ErrorCode.UnsupportedFormat, t('errors.odt.saveNeedsOriginal'))
  return original
}

export async function saveOdt(
  client: SidecarClient,
  sdocContent: string,
  target: { readonly origin: string | null },
): Promise<SavedDocx> {
  const original = assertOdtDestination(target.origin)
  const model = readSdocModel(sdocContent)

  const reply = await client.request(SidecarMethod.OdtSave, saveParamsOf(model, {}), new Uint8Array(original))
  const parsed = saveResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.odt.cannotSave'), t('errors.odt.saveContract'))
  }

  console.info(`[odt] preserved ${parsed.data.preservedBlocks}, rewritten ${parsed.data.rewrittenBlocks}`)
  return { bytes: reply.binary, inventory: parsed.data.inventory, original }
}
