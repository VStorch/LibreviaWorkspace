/**
 * An ODT opens as a file of its own: main keeps the original bytes, as for a `.docx`, and saving
 * goes back into that package through the sidecar, which rewrites only the edited blocks. A
 * document without an ODT original goes into the minimal package (`odt.create`), whole.
 */

import { walk } from '@services/document/export-common.js'
import { withoutUnsafeLinks } from '@services/document/link.js'
import { withEquationLatex } from '@services/document/mathml-latex.js'
import type { DocumentNode } from '@services/document/model.js'
import { AppError, ErrorCode } from '@shared/errors.js'
import {
  openResultSchema,
  readSdocModel,
  saveParamsOf,
  saveResultSchema,
  sdocContentOf,
  withoutOids,
  type OpenedDocx,
  type SavedDocx,
  type SdocModel,
} from '../docx/index.js'
import { readFileBytes } from '../fs/read-bytes.js'
import { isCreatedOriginal, OriginalFormat, originalOf, rememberOriginal } from '../opened-original.js'
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

export interface OdtTarget {
  /** `null` for a new document. Only compared with the stored original. */
  readonly origin: string | null
  /** `meta:generator` of a new package. */
  readonly generator: string
}

export interface SavedOdt extends SavedDocx {
  /** The starting package is the minimal one: the next save builds it again from the model. */
  readonly created: boolean
}

/**
 * Over the `.odt` the document came from, or the minimal package for any other: a new document, a
 * `.sdoc`, a DOCX. The minimal package is rebuilt on each save, so the page, sections, styles and
 * note numbering follow the model; another format's `oid`s mean nothing there.
 */
export async function saveOdt(
  client: SidecarClient,
  sdocContent: string,
  target: OdtTarget,
): Promise<SavedOdt> {
  const read = readSdocModel(sdocContent)
  const kept = isCreatedOriginal(target.origin) ? null : originalOf(target.origin, OriginalFormat.Odt)
  const model = kept === null ? { ...read, doc: withoutOids(read.doc) } : read
  const original = kept ?? Buffer.from(await createOdt(client, model, target.generator))
  const saved = await writeOdt(client, model, original)
  return { ...saved, original, created: kept === null }
}

/** "Export as": a new file from the model, without touching the open document's original. */
export async function exportOdt(
  client: SidecarClient,
  sdocContent: string,
  generator: string,
): Promise<Uint8Array> {
  const read = readSdocModel(sdocContent)
  const model = { ...read, doc: withoutOids(read.doc) }
  const saved = await writeOdt(client, model, Buffer.from(await createOdt(client, model, generator)))
  return saved.bytes
}

/** The equation's LaTeX goes along: ODF keeps it as the formula's description. */
async function writeOdt(
  client: SidecarClient,
  model: SdocModel,
  original: Buffer,
): Promise<Omit<SavedDocx, 'original'>> {
  const params = { ...saveParamsOf(model, {}), doc: withEquationLatex(model.doc as DocumentNode) }
  const reply = await client.request(SidecarMethod.OdtSave, params, new Uint8Array(original))
  const parsed = saveResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.odt.cannotSave'), t('errors.odt.saveContract'))
  }

  console.info(`[odt] preserved ${parsed.data.preservedBlocks}, rewritten ${parsed.data.rewrittenBlocks}`)
  return { bytes: reply.binary, inventory: parsed.data.inventory }
}

/** Created by the sidecar (`OdtTemplate`), with what lives in `styles.xml` and `meta.xml`. */
async function createOdt(client: SidecarClient, model: SdocModel, generator: string): Promise<Uint8Array> {
  const { page, sections, styles, notes, properties } = model
  const headingLevels = headingLevelsOf(model.doc as DocumentNode)
  const params = { page, sections, styles, notes, properties, headingLevels, generator }
  const reply = await client.request(SidecarMethod.OdtCreate, params)
  if (reply.binary.length === 0) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.odt.cannotCreate'), t('errors.odt.createContract'))
  }
  return reply.binary
}

/** A level without a style in the sheet gets the built-in heading in `styles.xml`. */
function headingLevelsOf(doc: DocumentNode): number[] {
  const levels = new Set<number>()
  walk(doc, (node) => {
    const level = Number(node.attrs?.['level'])
    if (node.type === 'heading' && Number.isInteger(level)) levels.add(level)
  })
  return [...levels]
}
