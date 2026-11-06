/**
 * The original bytes live in `opened-original.ts`, not in the stateless sidecar: its death does not
 * cost the surgical save.
 */

import { withoutOids } from '@services/document/oids.js'
import { AppError, ErrorCode } from '@shared/errors.js'
import { OriginalFormat, originalFormatOf, originalOf, rememberOriginal } from '../opened-original.js'
import {
  openResultSchema,
  readSdocModel,
  saveParamsOf,
  saveResultSchema,
  sdocContentOf,
  type OpenedPackage,
  type SavedPackage,
  type SdocModel,
} from '../package-document.js'
import type { SidecarClient } from '../sidecar/client.js'
import { SidecarMethod } from '../sidecar/protocol.js'
import { t } from '../i18n.js'
import { readFileBytes } from '../fs/read-bytes.js'

export async function openDocx(client: SidecarClient, path: string): Promise<OpenedPackage> {
  return decodeDocx(client, path, await readFileBytes(path))
}

/** Also for bytes that never reached the disk, like the personal data inspector's clean copy. */
export async function decodeDocx(client: SidecarClient, path: string, bytes: Buffer): Promise<OpenedPackage> {
  const reply = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
  const parsed = openResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.docx.cannotRead'), t('errors.docx.openContract'))
  }

  // Normalized, as `authorizePath` returns it and `origin` brings it: raw, the opened `.docx` would
  // be written over the minimal package.
  rememberOriginal(path, bytes, OriginalFormat.Docx)

  return { content: sdocContentOf(parsed.data.model), inventory: parsed.data.inventory }
}

export interface DocxTarget {
  /** `null` for a new document. Only compared with the stored original. */
  readonly origin: string | null
  readonly destination: string
  /**
   * The sidecar writes the template content type in the main part; without this, the document one.
   */
  readonly template?: boolean
}

/**
 * Over the original package, or the minimal package for a document born in the editor. The original
 * only applies if the edited document's origin is its path: otherwise a new document would inherit
 * another one's headers, styles and notes.
 */
export async function saveDocx(
  client: SidecarClient,
  sdocContent: string,
  target: DocxTarget,
): Promise<SavedPackage> {
  const model = modelForDocx(readSdocModel(sdocContent), target.origin)

  const kept = originalOf(target.origin, OriginalFormat.Docx)
  const original = kept ?? Buffer.from(await createDocx(client, model.page, model.styles))

  const reply = await client.request(
    SidecarMethod.DocxSave,
    saveParamsOf(model, target),
    new Uint8Array(original),
  )
  const parsed = saveResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.docx.cannotSave'), t('errors.docx.saveContract'))
  }

  console.info(`[docx] preserved ${parsed.data.preservedBlocks}, rewritten ${parsed.data.rewrittenBlocks}`)

  const inventory = parsed.data.inventory
  const lost = lossesOf(inventory.lost, model, kept === null)

  return {
    bytes: reply.binary,
    inventory: { ...inventory, lost },
    // The starting package of this save, not the saved bytes: starting from the previous result, a
    // new document would gain a copy of every image per save.
    original,
  }
}

/**
 * The `oid`s of another format's package mean nothing here: the document goes out whole from the
 * model, with no false "original package" loss.
 */
function modelForDocx(model: SdocModel, origin: string | null): SdocModel {
  const format = originalFormatOf(origin)
  return format === null || format === OriginalFormat.Docx ? model : { ...model, doc: withoutOids(model.doc) }
}

/** The sidecar's sentence is Portuguese; the screen shows the chosen language. */
function lossesOf(declared: readonly string[], model: SdocModel, fromMinimalPackage: boolean): string[] {
  const lost = [...declared]
  if (!fromMinimalPackage) return lost
  // The sidecar already declares the same loss when the relationship did not exist in the minimal
  // package: the sentence is a single one.
  if (
    [model.page, ...(Array.isArray(model.sections) ? model.sections : [])].some(hasForeignBands) &&
    !lost.includes(t('sidecar.notice.foreignBands'))
  )
    lost.push(t('sidecar.notice.foreignBands'))
  // Safety net: a model with `oid` was numbered against a package that is not here, and the minimal
  // package goes out. Said out loud, because silent loss is this program's worst defect.
  if (hasOid(model.doc)) lost.push(t('errors.docx.originPackage'))
  return lost
}

/** The `oid` is the block's fingerprint in the source package: one is enough. */
function hasOid(doc: unknown): boolean {
  if (typeof doc !== 'object' || doc === null) return false
  const content = (doc as { content?: unknown }).content
  if (!Array.isArray(content)) return false

  return content.some((block) => {
    if (typeof block !== 'object' || block === null) return false
    const attrs = (block as { attrs?: unknown }).attrs
    if (typeof attrs !== 'object' || attrs === null) return false
    return typeof (attrs as { oid?: unknown }).oid === 'string'
  })
}

const bandKeys = [
  'headerBand',
  'footerBand',
  'firstHeaderBand',
  'firstFooterBand',
  'evenHeaderBand',
  'evenFooterBand',
] as const

function hasForeignBands(page: unknown): boolean {
  if (typeof page !== 'object' || page === null) return false
  const record = page as Record<string, unknown>
  return bandKeys.some((key) => record[key] !== null && record[key] !== undefined)
}

/** Created by the sidecar (`DocxTemplate`), with the document's page and styles. */
async function createDocx(client: SidecarClient, page: unknown, styles: unknown): Promise<Uint8Array> {
  const reply = await client.request(SidecarMethod.DocxCreate, { page, styles })
  if (reply.binary.length === 0) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      t('errors.docx.cannotCreate'),
      t('errors.docx.createContract'),
    )
  }
  return reply.binary
}
