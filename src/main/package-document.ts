/** What DOCX and ODT share on their way to and from the sidecar: the same model, both ways. */

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
import { MAX_NAME_LENGTH } from '@shared/limits.js'
import { sidecarInventorySchema } from './sidecar/messages.js'
import { t } from './i18n.js'

/**
 * `page` and `doc` are validated in the renderer; the styles here, the only point they pass before
 * reaching the `.sdoc`.
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

export interface OpenedPackage {
  readonly content: string
  readonly inventory: LossInventory
}

export interface SavedPackage {
  readonly bytes: Uint8Array
  readonly inventory: LossInventory
  readonly original: Buffer
}

/** The constants, not a literal: a fixed `1` would push every opened document through migrations. */
export function sdocContentOf(model: object): string {
  return JSON.stringify({ format: SDOC_FORMAT, version: SDOC_VERSION, ...model })
}

export type SdocModel = ReturnType<typeof readSdocModel>

export function saveParamsOf(
  model: SdocModel,
  target: { readonly template?: boolean },
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
