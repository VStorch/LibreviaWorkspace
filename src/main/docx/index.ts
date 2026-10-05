/**
 * The original bytes live **here**, not in the stateless sidecar: its death does not cost the
 * surgical save.
 */

import { readFile } from 'node:fs/promises'
import { z } from 'zod'
import { SDOC_FORMAT, SDOC_VERSION } from '@services/document/serialize.js'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { Language, translate } from '@shared/i18n/index.js'
import {
  documentCommentSchema,
  documentNotesSchema,
  documentPropertiesSchema,
  styleSheetSchema,
} from '@shared/schemas.js'
import type { LossInventory } from '@shared/types.js'
import { normalizePath } from '../fs/paths.js'
import type { SidecarClient } from '../sidecar/client.js'
import { SidecarMethod } from '../sidecar/protocol.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'
import { MAX_NAME_LENGTH } from '@shared/limits.js'

/**
 * `ipc.ts` refuses more than 50 labels per category and more than 300 characters each: without the
 * cut, a pathological document would fail to open because of the warning.
 */
const inventoryLabels = z
  .array(z.string())
  .default([])
  .transform((labels) => labels.slice(0, 50).map((label) => label.slice(0, 300)))

const inventorySchema = z.object({
  invisible: inventoryLabels,
  lost: inventoryLabels,
  structural: inventoryLabels,
})

/**
 * `page` and `doc` are validated in the renderer; the styles here, the only point they pass before
 * reaching the `.sdoc`.
 */
const openResultSchema = z.object({
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
  inventory: inventorySchema,
})

const saveResultSchema = z.object({
  inventory: inventorySchema,
  preservedBlocks: z.number().int().nonnegative(),
  rewrittenBlocks: z.number().int().nonnegative(),
})

/** As they were on disk when opened, not as they are now, which someone else may have changed. */
let openedOriginal: { path: string; bytes: Buffer } | null = null

export function forgetOpenedDocx(): void {
  openedOriginal = null
}

/**
 * After a crash the original bytes died with the process: without rereading them, saving would be
 * refused.
 */
export async function adoptDocxOriginal(path: string): Promise<boolean> {
  try {
    openedOriginal = { path: normalizePath(path), bytes: await readFile(path) }
    return true
  } catch {
    openedOriginal = null
    return false
  }
}

export interface OpenedDocx {
  readonly content: string
  readonly inventory: LossInventory
}

export async function openDocx(client: SidecarClient, path: string): Promise<OpenedDocx> {
  let bytes: Buffer
  try {
    bytes = await readFile(path)
  } catch (cause) {
    throw fromFileSystemError(cause, 'leitura', editorPreferences().language)
  }

  const reply = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
  const parsed = openResultSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(ErrorCode.SidecarFailed, t('errors.docx.cannotRead'), t('errors.docx.openContract'))
  }

  // Normalized, as `authorizePath` returns it and `origin` brings it: raw, the opened `.docx` would
  // be written over the minimal package.
  openedOriginal = { path: normalizePath(path), bytes }

  return {
    // The constants, not a literal: a fixed `1` would push every Word document through migrations.
    content: JSON.stringify({ format: SDOC_FORMAT, version: SDOC_VERSION, ...parsed.data.model }),
    inventory: parsed.data.inventory,
  }
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
  const model = unwrapSdoc(sdocContent)

  const kept =
    openedOriginal !== null && target.origin !== null && openedOriginal.path === normalizePath(target.origin)
      ? openedOriginal.bytes
      : null
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

  console.info(`[docx] preservados ${parsed.data.preservedBlocks}, reescritos ${parsed.data.rewrittenBlocks}`)

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

type SdocModel = ReturnType<typeof unwrapSdoc>

function saveParamsOf(model: SdocModel, target: DocxTarget): Record<string, unknown> {
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
    // The sidecar only writes note numbering and properties when they differ from the target file.
    ...(model.notes === undefined ? {} : { notes: model.notes }),
    ...(model.properties === undefined ? {} : { properties: model.properties }),
    ...(model.styles === undefined ? {} : { styles: model.styles }),
    ...(target.template === true ? { template: true } : {}),
  }
}

/** The sidecar's sentence is Portuguese; the screen shows the chosen language. */
function lossesOf(declared: readonly string[], model: SdocModel, fromMinimalPackage: boolean): string[] {
  const foreignBandsPt = translate(Language.Portuguese, 'errors.docx.foreignBands')
  const lost = declared.map((item) => (item === foreignBandsPt ? t('errors.docx.foreignBands') : item))
  if (!fromMinimalPackage) return lost
  // The sidecar already declares the same loss when the relationship did not exist in the minimal
  // package: the sentence is a single one.
  if (
    [model.page, ...(Array.isArray(model.sections) ? model.sections : [])].some(hasForeignBands) &&
    !lost.includes(t('errors.docx.foreignBands'))
  )
    lost.push(t('errors.docx.foreignBands'))
  // Safety net: a model with `oid` was numbered against a package that is not here, and the minimal
  // package goes out. Said out loud, because silent loss is this program's worst defect.
  if (hasOid(model.doc)) lost.push(t('errors.docx.originPackage'))
  return lost
}

/**
 * Called **after** the save reached the disk, so a failure does not leave the original pointing to
 * a path the document does not have. Applies to any destination: a `.docx` saved as `.sdoc` keeps
 * the same `oid`s.
 */
export function followDocxOriginal(
  origin: string | null,
  destination: string,
  saved: SavedDocx | null,
): void {
  const target = normalizePath(destination)
  if (saved !== null) {
    openedOriginal = { path: target, bytes: saved.original }
    return
  }
  if (openedOriginal !== null && origin !== null && openedOriginal.path === normalizePath(origin)) {
    openedOriginal = { path: target, bytes: openedOriginal.bytes }
  }
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

function unwrapSdoc(content: string): {
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
  notes: unknown
  properties: unknown
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
      notes: z.unknown().optional(),
      // Checked here too: they go into the user's file.
      properties: documentPropertiesSchema.optional(),
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
    notes: envelope.data.notes,
    properties: envelope.data.properties,
  }
}
