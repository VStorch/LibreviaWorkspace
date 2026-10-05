import { z } from 'zod'
import { AppError, ErrorCode } from '@shared/errors.js'
import { Language, translate } from '@shared/i18n/index.js'
import {
  documentCommentSchema,
  documentNotesSchema,
  documentPropertiesSchema,
  pageSetupSchema,
  sectionSetupSchema,
  styleSheetSchema,
} from '@shared/schemas.js'
import {
  DEFAULT_PAGE_SETUP,
  isValidMargins,
  type DocumentComment,
  type DocumentModel,
  type DocumentNode,
  type DocumentNotes,
  type DocumentProperties,
  type NoteNumbering,
} from './model.js'
import { LEGACY_STYLES, type StyleSheet } from './styles.js'

/**
 * The internal `.sdoc` format.
 *
 * It is JSON: the document model saved as is. It does not replace DOCX; it saves and reopens **with
 * no loss at all**, which `.txt` cannot. Images are embedded as data URIs; if that gets heavy, the
 * container can become a ZIP without anything outside this file changing.
 *
 * `version` keeps a file saved today readable as the model evolves. Each version that changes the
 * document shape gets a migration in `migrate`, applied on read:
 *
 * - **2**: the image stopped being a block and moved inside the paragraph, as in Word. Editing it
 *   as a block split the surrounding paragraph.
 * - **3**: the document carries its **styles** (`styles.ts`). A version 2 file has none and gets
 *   `LEGACY_STYLES` on read: the look the editor already drew, measure by measure, so the old
 *   document opens identical.
 * - **4**: a block carries only **direct** formatting; the inherited part comes from the styles.
 *   Blocks from an older file stay flattened, and reading marks them (`flattened`) so saving to
 *   DOCX compares them with a flattened reading of the original. The nodes are untouched:
 *   unflattening would need each one's `styles.xml`, and the flattened form draws the same.
 * - **5**: the `.docx` reader produces bookmarks, fields, internal links and tables of contents. An
 *   older draft lacks them in the nodes, and reading marks it (`beforeReferences`) for the same
 *   reason as version 4.
 * - **6**: the document has **sections**: `sections` holds the ones before the last, and the
 *   paragraph closing each one carries `sectionBreak`. An older draft has neither (its page is the
 *   whole document's), and reading marks it (`beforeSections`) for the same reason as version 4.
 * - **7**: the document has **comments**: `comments` holds each body, and the text holds the anchor
 *   ends (`commentStart` and `commentEnd`). An older draft has no ends, and reading marks it
 *   (`beforeComments`) for the same reason as version 4.
 * - **8**: the text carries **revisions**: the `insertion` and `deletion` marks, the paragraph mark
 *   revision (`markRevision`) and the table row one (`rowRevision`); `trackChanges` is the document
 *   switch. An older draft has none (insertions were plain text and deletions did not show), and
 *   reading marks it (`beforeRevisions`) for the same reason as version 4.
 * - **9**: the text carries footnotes and endnotes: the reference is the `noteRef` node, with the
 *   note body inside, and `notes` holds the document numbering. An older draft has no references,
 *   and reading marks it (`beforeNotes`) for the same reason as version 4.
 * - **10**: the document **properties** (title, subject, author…) in `properties`. An older draft
 *   simply lacks them and gets no flag: when saving to DOCX their absence means "leave the file's
 *   as they are".
 * - **11**: **equations**: the `math` node, with the file's OMML inside and the MathML the screen
 *   draws. An older draft has no node (the equation was hidden in the paragraph), and reading marks
 *   it (`beforeMath`) for the same reason as version 4.
 */
export const SDOC_FORMAT = 'sdoc'
export const SDOC_VERSION = 11

/** Content is only validated in shape; the fine structure is ProseMirror's. */
const documentNodeSchema: z.ZodType<DocumentNode> = z.looseObject({
  type: z.string(),
})

const sdocSchema = z.object({
  format: z.literal(SDOC_FORMAT),
  version: z.number().int().positive(),
  page: pageSetupSchema,
  doc: documentNodeSchema,
  // Optional because version 2 has no styles: `migrate` decides what to do with their absence, not
  // the schema.
  styles: styleSheetSchema.optional(),
  // Only present when true; see `DocumentModel.flattened`.
  flattened: z.boolean().optional(),
  // Only present when true; see `DocumentModel.beforeReferences`.
  beforeReferences: z.boolean().optional(),
  // See `DocumentModel.sections` and `beforeSections`.
  sections: z.array(sectionSetupSchema).max(10_000).optional(),
  beforeSections: z.boolean().optional(),
  outsideBookmarks: z.array(z.string()).optional(),
  // See `DocumentModel.comments` and `beforeComments`.
  comments: z.array(documentCommentSchema).max(100_000).optional(),
  beforeComments: z.boolean().optional(),
  // See `DocumentModel.trackChanges` and `beforeRevisions`.
  trackChanges: z.boolean().optional(),
  beforeRevisions: z.boolean().optional(),
  // See `DocumentModel.notes` and `beforeNotes`.
  notes: documentNotesSchema.optional(),
  beforeNotes: z.boolean().optional(),
  // See `DocumentModel.beforeMath`.
  beforeMath: z.boolean().optional(),
  // See `DocumentModel.properties`.
  properties: documentPropertiesSchema.optional(),
})

export function serializeDocument(model: DocumentModel): string {
  return JSON.stringify(
    {
      format: SDOC_FORMAT,
      version: SDOC_VERSION,
      page: model.page,
      doc: model.doc,
      // In the envelope, as the sidecar puts them when opening a `.docx`: the nodes, and their
      // fingerprints, stay as they were.
      styles: model.styles,
      ...onlyTrue('flattened', model.flattened),
      ...onlyTrue('beforeReferences', model.beforeReferences),
      ...onlyNonEmpty('sections', model.sections),
      ...onlyTrue('beforeSections', model.beforeSections),
      ...onlyDefined('outsideBookmarks', model.outsideBookmarks),
      ...onlyNonEmpty('comments', model.comments),
      ...onlyTrue('beforeComments', model.beforeComments),
      ...onlyDefined('trackChanges', model.trackChanges),
      ...onlyTrue('beforeRevisions', model.beforeRevisions),
      ...onlyDefined('notes', model.notes),
      ...onlyTrue('beforeNotes', model.beforeNotes),
      ...onlyTrue('beforeMath', model.beforeMath),
      ...onlyDefined('properties', model.properties),
    },
    null,
    2,
  )
}

/** Old-document flags only exist when on: a new file does not carry them. */
export function onlyTrue<K extends string>(key: K, value: boolean | undefined): Partial<Record<K, true>> {
  return value === true ? ({ [key]: true } as Record<K, true>) : {}
}

export function onlyDefined<K extends string, V>(key: K, value: V | undefined): Partial<Record<K, V>> {
  return value === undefined ? {} : ({ [key]: value } as Record<K, V>)
}

export function onlyNonEmpty<K extends string, V>(
  key: K,
  value: readonly V[] | undefined,
): Partial<Record<K, V[]>> {
  return value === undefined || value.length === 0 ? {} : ({ [key]: [...value] } as Record<K, V[]>)
}

/**
 * A corrupt file or one from a future version gives a sentence the user understands, not a JSON
 * error.
 */
export function parseDocument(text: string, language: Language = Language.Portuguese): DocumentModel {
  let raw: unknown
  try {
    raw = JSON.parse(text)
  } catch {
    throw new AppError(ErrorCode.UnsupportedFormat, translate(language, 'errors.document.corrupt'))
  }

  const parsed = sdocSchema.safeParse(raw)
  if (!parsed.success) {
    throw new AppError(ErrorCode.UnsupportedFormat, translate(language, 'errors.document.invalid'))
  }

  if (parsed.data.version > SDOC_VERSION) {
    throw new AppError(ErrorCode.UnsupportedFormat, translate(language, 'errors.document.newerVersion'))
  }

  // Invalid margins do not stop reading: the document is recovered with the default setup, because
  // the user's text is worth more than the layout.
  const page = isValidMargins(parsed.data.page) ? parsed.data.page : DEFAULT_PAGE_SETUP
  const sections = (parsed.data.sections ?? []).map((section) =>
    isValidMargins(section) ? section : { ...DEFAULT_PAGE_SETUP, id: section.id },
  )

  return {
    page,
    doc: migrate(parsed.data.doc, parsed.data.version),
    styles: migrateStyles(parsed.data.styles, parsed.data.version),
    ...legacyFlagsOf(parsed.data),
    ...(sections.length > 0 ? { sections } : {}),
    ...optionalPartsOf(parsed.data),
  }
}

type SdocData = z.infer<typeof sdocSchema>

/** Each old-file flag and the `.sdoc` version that brought the feature. */
const LEGACY_FLAGS = [
  ['flattened', 4],
  ['beforeReferences', 5],
  ['beforeSections', 6],
  ['beforeComments', 7],
  ['beforeRevisions', 8],
  ['beforeNotes', 9],
  ['beforeMath', 11],
] as const

type LegacyFlag = (typeof LEGACY_FLAGS)[number][0]

function legacyFlagsOf(data: SdocData): Partial<Pick<DocumentModel, LegacyFlag>> {
  return Object.fromEntries(
    LEGACY_FLAGS.filter(([flag, since]) => data.version < since || data[flag] === true).map(([flag]) => [
      flag,
      true,
    ]),
  )
}

function optionalPartsOf(
  data: SdocData,
): Pick<DocumentModel, 'outsideBookmarks' | 'comments' | 'trackChanges' | 'notes' | 'properties'> {
  return {
    ...(data.outsideBookmarks === undefined ? {} : { outsideBookmarks: data.outsideBookmarks }),
    ...(data.comments === undefined ? {} : { comments: data.comments.map(commentOf) }),
    ...(data.trackChanges === undefined ? {} : { trackChanges: data.trackChanges }),
    ...(data.notes === undefined ? {} : { notes: notesOf(data.notes) }),
    ...(data.properties === undefined ? {} : { properties: propertiesOf(data.properties) }),
  }
}

/** Without absent keys, for `exactOptionalPropertyTypes`. */
export function propertiesOf(raw: z.infer<typeof documentPropertiesSchema>): DocumentProperties {
  return Object.fromEntries(
    Object.entries(raw).filter(([, value]) => value !== undefined),
  ) as DocumentProperties
}

/** Without absent keys, for `exactOptionalPropertyTypes`. */
export function notesOf(raw: z.infer<typeof documentNotesSchema>): DocumentNotes {
  const numbering = (pr: NonNullable<typeof raw.footnotePr>): NoteNumbering => ({
    ...(pr.numFmt === undefined ? {} : { numFmt: pr.numFmt }),
    ...(pr.start === undefined ? {} : { start: pr.start }),
    ...(pr.restart === undefined ? {} : { restart: pr.restart }),
    ...(pr.pos === undefined ? {} : { pos: pr.pos }),
  })
  return {
    ...(raw.footnotePr === undefined ? {} : { footnotePr: numbering(raw.footnotePr) }),
    ...(raw.endnotePr === undefined ? {} : { endnotePr: numbering(raw.endnotePr) }),
  }
}

/** Without absent keys, for `exactOptionalPropertyTypes`. */
export function commentOf(raw: z.infer<typeof documentCommentSchema>): DocumentComment {
  return {
    id: raw.id,
    author: raw.author,
    date: raw.date,
    paragraphs: raw.paragraphs,
    done: raw.done,
    ...(raw.parentId === undefined ? {} : { parentId: raw.parentId }),
    ...(raw.initials === undefined ? {} : { initials: raw.initials }),
    ...(raw.paraId === undefined ? {} : { paraId: raw.paraId }),
    ...(raw.rich === true ? { rich: true } : {}),
  }
}

function migrate(doc: DocumentNode, version: number): DocumentNode {
  return version < 2 ? wrapLooseImages(doc) : doc
}

/**
 * `LEGACY_STYLES`, not `BUILTIN_STYLES`: the old file must reopen with its old pagination. By
 * version, not by the field's presence: a version 2 file with `styles` is a patched file.
 */
function migrateStyles(styles: StyleSheet | undefined, version: number): StyleSheet {
  return version < 3 || styles === undefined ? LEGACY_STYLES : styles
}

/** Inline images only belong in text nodes. */
const TEXTBLOCKS = new Set(['paragraph', 'heading', 'codeBlock'])

/**
 * Wraps in a paragraph each image version 1 left loose between blocks. Tiptap builds content
 * without validating, and without this the image would survive by chance until the first path that
 * validates.
 */
function wrapLooseImages(node: DocumentNode): DocumentNode {
  if (node.content === undefined || TEXTBLOCKS.has(node.type)) return node

  return {
    ...node,
    content: node.content.map((child) =>
      child.type === 'image' ? { type: 'paragraph', content: [child] } : wrapLooseImages(child),
    ),
  }
}

export function isDocumentFile(text: string): boolean {
  return text.trimStart().startsWith('{') && text.includes(`"${SDOC_FORMAT}"`)
}
