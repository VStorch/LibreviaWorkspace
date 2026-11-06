/**
 * The original bytes live in `opened-original.ts`, not in the stateless sidecar: its death does not
 * cost the surgical save.
 */

import { z } from 'zod'
import { SDOC_FORMAT, SDOC_VERSION } from '@services/document/serialize.js'
import { AppError, ErrorCode } from '@shared/errors.js'
import {
  documentCommentSchema,
  documentNotesSchema,
  documentPropertiesSchema,
  documentStatisticsSchema,
  styleSheetSchema,
} from '@shared/schemas.js'
import type { LossInventory } from '@shared/types.js'
import {
  adoptOriginal,
  followOriginal,
  forgetOpenedOriginal,
  OriginalFormat,
  originalFormatOf,
  originalOf,
  rememberOriginal,
} from '../opened-original.js'
import type { SidecarClient } from '../sidecar/client.js'
import { sidecarInventorySchema } from '../sidecar/messages.js'
import { SidecarMethod } from '../sidecar/protocol.js'
import { t } from '../i18n.js'
import { readFileBytes } from '../fs/read-bytes.js'
import { MAX_NAME_LENGTH } from '@shared/limits.js'

/**
 * `page` and `doc` are validated in the renderer; the styles here, the only point they pass before
 * reaching the `.sdoc`. The ODT import answers in the same shape.
 */
export const openResultSchema = z.object({
  model: z.object({
    page: z.unknown(),
    sections: z.unknown().optional(),
    doc: z.unknown(),
    styles: styleSheetSchema,
    outsideBookmarks: z.array(z.string().max(MAX_NAME_LENGTH)).max(10_000).optional(),
    comments: z.array(documentCommentSchema).max(100_000).optional(),
    trackChanges: z.boolean().optional(),
    notes: documentNotesSchema.optional(),
    properties: documentPropertiesSchema.optional(),
  }),
  inventory: sidecarInventorySchema,
})

export const saveResultSchema = z.object({
  inventory: sidecarInventorySchema,
  preservedBlocks: z.number().int().nonnegative(),
  rewrittenBlocks: z.number().int().nonnegative(),
})

export function forgetOpenedDocx(): void {
  forgetOpenedOriginal()
}

/** See `adoptOriginal`. */
export async function adoptDocxOriginal(path: string): Promise<boolean> {
  return adoptOriginal(path, OriginalFormat.Docx)
}

export interface OpenedDocx {
  readonly content: string
  readonly inventory: LossInventory
}

export async function openDocx(client: SidecarClient, path: string): Promise<OpenedDocx> {
  return decodeDocx(client, path, await readFileBytes(path))
}

/** Also for bytes that never reached the disk, like the personal data inspector's clean copy. */
export async function decodeDocx(client: SidecarClient, path: string, bytes: Buffer): Promise<OpenedDocx> {
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

/** The constants, not a literal: a fixed `1` would push every opened document through migrations. */
export function sdocContentOf(model: object): string {
  return JSON.stringify({ format: SDOC_FORMAT, version: SDOC_VERSION, ...model })
}

export interface SavedDocx {
  readonly bytes: Uint8Array
  readonly inventory: LossInventory
  readonly original: Buffer
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
): Promise<SavedDocx> {
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

export type SdocModel = ReturnType<typeof readSdocModel>

/**
 * The `oid`s of another format's package mean nothing here: the document goes out whole from the
 * model, with no false "original package" loss.
 */
function modelForDocx(model: SdocModel, origin: string | null): SdocModel {
  const format = originalFormatOf(origin)
  return format === null || format === OriginalFormat.Docx ? model : { ...model, doc: withoutOids(model.doc) }
}

export function saveParamsOf(
  model: SdocModel,
  target: Pick<DocxTarget, 'template'>,
): Record<string, unknown> {
  // `flatten` picks the sidecar's reference reading for the flattened draft. Styles always go, to
  // reach `word/styles.xml`.
  return {
    page: model.page,
    doc: model.doc,
    ...(model.flattened ? { flatten: true } : {}),
    ...(model.beforeReferences ? { beforeReferences: true } : {}),
    ...(model.sections === undefined ? {} : { sections: model.sections }),
    ...(model.beforeSections ? { beforeSections: true } : {}),
    // The empty list means "all deleted"; for a draft older than comments, only the flag.
    ...(model.beforeComments ? { beforeComments: true } : { comments: model.comments }),
    ...(model.trackChanges === undefined ? {} : { trackChanges: model.trackChanges }),
    ...(model.beforeRevisions ? { beforeRevisions: true } : {}),
    ...(model.beforeNotes ? { beforeNotes: true } : {}),
    ...(model.beforeMath ? { beforeMath: true } : {}),
    ...(model.beforeVerticalMerges ? { beforeVerticalMerges: true } : {}),
    // The sidecar only writes note numbering and properties when they differ from the target file.
    ...(model.notes === undefined ? {} : { notes: model.notes }),
    ...propertiesParam(model),
    ...(model.styles === undefined ? {} : { styles: model.styles }),
    ...(target.template === true ? { template: true } : {}),
  }
}

/** The counts travel with the properties: the sidecar writes both to `docProps/`. */
function propertiesParam(model: SdocModel): Record<string, unknown> {
  if (model.properties === undefined && model.statistics === undefined) return {}
  return {
    properties: {
      ...(model.properties ?? {}),
      ...(model.statistics === undefined ? {} : { statistics: model.statistics }),
    },
  }
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

/** See `followOriginal`. */
export function followDocxOriginal(
  origin: string | null,
  destination: string,
  saved: SavedDocx | null,
): void {
  followOriginal(
    origin,
    destination,
    saved === null ? null : { original: saved.original, format: OriginalFormat.Docx },
  )
}

/** At any depth: a DOCX list item carries its own. */
export function withoutOids(node: unknown): unknown {
  if (Array.isArray(node)) return node.map(withoutOids)
  if (typeof node !== 'object' || node === null) return node
  const record = node as Record<string, unknown>
  const { attrs, content } = record
  const copy: Record<string, unknown> = { ...record }
  if (typeof attrs === 'object' && attrs !== null && 'oid' in attrs) {
    const rest = Object.entries(attrs).filter(([name]) => name !== 'oid')
    if (rest.length > 0) copy['attrs'] = Object.fromEntries(rest)
    else delete copy['attrs']
  }
  if (Array.isArray(content)) copy['content'] = content.map(withoutOids)
  return copy
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

export function readSdocModel(content: string): {
  page: unknown
  doc: unknown
  styles: unknown
  flattened: boolean
  beforeReferences: boolean
  sections: unknown
  beforeSections: boolean
  beforeComments: boolean
  comments: unknown[]
  trackChanges: boolean | undefined
  beforeRevisions: boolean
  beforeNotes: boolean
  beforeMath: boolean
  beforeVerticalMerges: boolean
  notes: unknown
  properties: unknown
  statistics: unknown
} {
  let parsed: unknown
  try {
    parsed = JSON.parse(content)
  } catch {
    throw new AppError(ErrorCode.Internal, t('errors.docx.inconsistentState'))
  }

  // The sidecar only touches `word/styles.xml` when some style changed.
  const envelope = z
    .object({
      page: z.unknown(),
      doc: z.unknown(),
      styles: styleSheetSchema.optional(),
      flattened: z.boolean().optional(),
      beforeReferences: z.boolean().optional(),
      sections: z.unknown().optional(),
      beforeSections: z.boolean().optional(),
      beforeComments: z.boolean().optional(),
      comments: z.array(documentCommentSchema).max(100_000).optional(),
      trackChanges: z.boolean().optional(),
      beforeRevisions: z.boolean().optional(),
      beforeNotes: z.boolean().optional(),
      beforeMath: z.boolean().optional(),
      beforeVerticalMerges: z.boolean().optional(),
      notes: z.unknown().optional(),
      // Checked here too: they go into the user's file.
      properties: documentPropertiesSchema.optional(),
      statistics: documentStatisticsSchema.optional(),
    })
    .safeParse(parsed)
  if (!envelope.success) {
    throw new AppError(ErrorCode.Internal, t('errors.docx.inconsistentState'))
  }

  return {
    page: envelope.data.page,
    doc: envelope.data.doc,
    styles: envelope.data.styles,
    flattened: envelope.data.flattened === true,
    beforeReferences: envelope.data.beforeReferences === true,
    sections: envelope.data.sections,
    beforeSections: envelope.data.beforeSections === true,
    beforeComments: envelope.data.beforeComments === true,
    comments: envelope.data.comments ?? [],
    trackChanges: envelope.data.trackChanges,
    beforeRevisions: envelope.data.beforeRevisions === true,
    beforeNotes: envelope.data.beforeNotes === true,
    beforeMath: envelope.data.beforeMath === true,
    beforeVerticalMerges: envelope.data.beforeVerticalMerges === true,
    notes: envelope.data.notes,
    properties: envelope.data.properties,
    statistics: envelope.data.statistics,
  }
}
