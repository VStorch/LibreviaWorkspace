import { z } from 'zod'
import type { DocumentNode } from '@services/document/model.js'
import { MAX_ZOOM, MIN_ZOOM } from '@services/document/zoom.js'
import { PRIVACY_CATEGORIES } from '@services/file/privacy.js'
import { CUSTOM_PROPERTY_TYPES } from '@services/file/properties.js'
import { Language, LANGUAGES } from './i18n/language.js'
import { ACP_AGENT_IDS } from '@services/ai/acp-agents.js'
import { AiProposalKind } from '@services/ai/proposals.js'
import { AiProviderKind, AiRole } from '@services/ai/provider.js'
import {
  AiAction,
  AiFormat,
  AiSheetTask,
  MCP_DOCUMENT_REFUSED,
  McpDocumentAsk,
  McpDocumentRefusal,
  Theme,
} from './types.js'
import {
  MAX_AI_AGENT_ARGS_LENGTH,
  MAX_AI_COMMENT_LENGTH,
  MAX_AI_FIND_LENGTH,
  MAX_AI_PARAGRAPHS,
  MAX_AI_REASON_LENGTH,
  MAX_AI_REPLACE_LENGTH,
  MAX_AI_CHAT_MESSAGE_LENGTH,
  MAX_AI_CHAT_MESSAGES,
  MAX_AI_INSTRUCTION_LENGTH,
  MAX_AI_KEY_LENGTH,
  MAX_AI_RESPONSE_LENGTH,
  MAX_AI_SELECTION_LENGTH,
  MAX_AI_URL_LENGTH,
  MAX_AI_DOCUMENT_LENGTH,
  MAX_AI_FILL_ROWS,
  MAX_AI_FORMULA_LENGTH,
  MAX_AI_SHEET_CONTEXT_LENGTH,
  MAX_MCP_CLIENT_NAME_LENGTH,
  MAX_MCP_CLIENTS,
  MAX_COLOR_LENGTH,
  MAX_CSS_VALUE_LENGTH,
  MAX_DESCRIPTION_LENGTH,
  MAX_FIELD_LENGTH,
  MAX_CUSTOM_PROPERTIES,
  MAX_ID_LENGTH,
  MAX_IMAGE_SIDE_PX,
  MAX_NAME_LENGTH,
  MAX_PATH_LENGTH,
  MAX_PROPERTY_LENGTH,
  MAX_SECTION_COLUMNS,
  MAX_START_NUMBER,
} from './limits.js'

/**
 * Shared schemas. Page setup is validated both when reading a `.sdoc` from disk and when receiving
 * a print request from the renderer; two definitions would drift, and the drift would show as a
 * wrong margin on paper.
 */

export const bandPieceSchema = z.object({
  kind: z.enum(['text', 'image', 'pageNumber', 'totalPages']),
  text: z.string().max(1000).optional(),
  /** A header image is small: a logo, not a photo. */
  src: z.string().max(4_000_000).optional(),
  width: z.number().int().positive().max(MAX_IMAGE_SIDE_PX).optional(),
  height: z.number().int().positive().max(MAX_IMAGE_SIDE_PX).optional(),
  bold: z.boolean().default(false),
  italic: z.boolean().default(false),
  color: z.string().max(MAX_COLOR_LENGTH).optional(),
  fontSize: z.string().max(MAX_CSS_VALUE_LENGTH).optional(),
  /** A CSS font stack, as the reader resolved it. */
  fontFamily: z.string().max(MAX_NAME_LENGTH).optional(),
  /** The piece starts a new line: in the file it begins another paragraph. */
  line: z.boolean().optional(),
  /**
   * Where the piece lives in the file: the relationship, the paragraph and the piece in it.
   *
   * Without declaring the field, zod would **drop it silently**, and text typed in the header would
   * return to the screen but not to the `.docx`.
   */
  pid: z.string().max(MAX_ID_LENGTH).optional(),
  /** The file has `{n}` or `{total}` written as text, not as a field. */
  literal: z.boolean().optional(),
})

/**
 * Checked only as far as "is a node", on purpose. The real validation is the ProseMirror
 * serializer, which builds from the editor schema and **only emits what it knows**: that is the
 * barrier that keeps a document from injecting markup into the page. Repeating the node type list
 * here would duplicate it where it cannot be checked against the editor, and the older copy would
 * win.
 */
const documentNodeSchema = z.custom<DocumentNode>(
  (value) => typeof value === 'object' && value !== null && typeof (value as DocumentNode).type === 'string',
)

/**
 * Left open on purpose: the geometry comes from the file and `services/document/floating.ts`
 * interprets it. Validating field by field here would duplicate that interpretation.
 */
const bandFloatSchema = z.object({
  kind: z.enum(['image', 'text', 'rule']),
  src: z.string().optional(),
  /**
   * Open like the rest of this schema, for the same reason: the editor serializer interprets it and
   * only emits what its schema knows. Without the field, zod would **drop it silently**, and the
   * header title box would appear on the sheet at the right size and empty.
   */
  content: z.array(documentNodeSchema).max(200).optional(),
  /** Where the box lives in the file, when its text is editable. */
  bid: z.string().max(MAX_ID_LENGTH).optional(),
  /** Border and fill, when the reader could reproduce them. */
  fill: z.string().max(MAX_COLOR_LENGTH).optional(),
  line: z.string().max(MAX_COLOR_LENGTH).optional(),
  lineWidthPt: z.number().min(0).max(200).optional(),
  dash: z.boolean().optional(),
  widthMm: z.number(),
  heightMm: z.number(),
  rotation: z.number(),
  hFrom: z.string(),
  hOffsetMm: z.number().optional(),
  hAlign: z.string().optional(),
  vFrom: z.string(),
  vOffsetMm: z.number().optional(),
  vAlign: z.string().optional(),
  behind: z.boolean(),
  wrap: z.string(),
  dxMm: z.number().optional(),
  dyMm: z.number().optional(),
})

/**
 * `borders` lists the initials of the sides with a line (`t`, `l`, `b`, `r`), already resolved by
 * the reader: in OOXML each side comes through three paths, and redoing that in two renderers is
 * how screen and paper drift apart.
 */
const bandCellSchema = z.object({
  pieces: z.array(bandPieceSchema).max(40).default([]),
  width: z.number().min(0).max(1).default(0),
  span: z.number().int().min(1).max(32).default(1),
  rowSpan: z.number().int().min(1).max(32).default(1),
  align: z.string().max(MAX_CSS_VALUE_LENGTH).optional(),
  borders: z.string().max(4).default(''),
})

/**
 * A header or footer from a Word document. Three columns and a rule, the layout Word always used,
 * which covers almost every corporate header. Pieces with an address are editable; the rest of the
 * OOXML part goes back to the file untouched.
 */
export const bandSchema = z.object({
  left: z.array(bandPieceSchema).max(20).default([]),
  center: z.array(bandPieceSchema).max(20).default([]),
  right: z.array(bandPieceSchema).max(20).default([]),
  rule: z.boolean().default(false),
  floats: z.array(bandFloatSchema).max(20).default([]),
  rows: z
    .array(z.object({ cells: z.array(bandCellSchema).max(32).default([]) }))
    .max(32)
    .default([]),
})

export const pageSetupSchema = z.object({
  size: z.enum(['A4', 'Letter']),
  orientation: z.enum(['portrait', 'landscape']),
  margins: z.object({
    top: z.number(),
    right: z.number(),
    bottom: z.number(),
    left: z.number(),
  }),
  // Optional so documents saved without them keep opening: a compatible addition needs no new
  // format version.
  header: z.string().max(500).default(''),
  footer: z.string().max(500).default(''),
  // Optional for the same reason. When present they win on screen: they are the document's real
  // header, and the text above is what the user typed in a document created here.
  headerBand: bandSchema.nullable().default(null),
  footerBand: bandSchema.nullable().default(null),
  // First page and even pages, when the document asks for them. Optional for the same reason.
  firstHeaderBand: bandSchema.nullable().default(null),
  firstFooterBand: bandSchema.nullable().default(null),
  evenHeaderBand: bandSchema.nullable().default(null),
  evenFooterBand: bandSchema.nullable().default(null),
  /**
   * Vertical origin for anchors inside the header: they are relative to the paragraph, and the
   * header paragraph starts here.
   */
  headerDistanceMm: z.number().default(12.5),
  footerDistanceMm: z.number().default(12.5),
  // Page numbering and band switches. Optional for the same reason: when absent, saving leaves what
  // the file says.
  pageNumberFormat: z.enum(['decimal', 'lowerRoman', 'upperRoman', 'lowerLetter', 'upperLetter']).optional(),
  pageNumberStart: z.number().int().min(0).max(MAX_START_NUMBER).nullable().optional(),
  titlePage: z.boolean().nullable().optional(),
  evenAndOddHeaders: z.boolean().nullable().optional(),
  // Optional for the same reason.
  start: z.enum(['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn']).optional(),
  columns: z
    .object({
      count: z.number().int().min(1).max(MAX_SECTION_COLUMNS),
      spaceMm: z.number().min(0).max(1000),
      separator: z.boolean(),
      widthsMm: z.array(z.number()).max(MAX_SECTION_COLUMNS).optional(),
    })
    .optional(),
})

/** A section before the last one: its setup and the id of the mark that closes it. */
export const sectionSetupSchema = pageSetupSchema.extend({
  id: z.string().min(1).max(MAX_FIELD_LENGTH),
})

/**
 * See `DocumentComment`. Validated on input, like the styles: it comes from someone else's file and
 * ends up in the user's `.sdoc`. The caps only stop pathological files.
 */
export const documentCommentSchema = z.object({
  id: z.string().min(1).max(MAX_FIELD_LENGTH),
  parentId: z.string().max(MAX_FIELD_LENGTH).optional(),
  author: z.string().max(1000),
  initials: z.string().max(MAX_FIELD_LENGTH).optional(),
  date: z.string().max(MAX_FIELD_LENGTH),
  paragraphs: z.array(z.string().max(100_000)).max(1000),
  done: z.boolean(),
  paraId: z.string().max(MAX_FIELD_LENGTH).optional(),
  rich: z.boolean().optional(),
})

/**
 * `w:footnotePr`/`w:endnotePr`. Goes into the `.sdoc`, so it is checked on input, like comments.
 */
const notePrSchema = z.object({
  numFmt: z.string().max(MAX_FIELD_LENGTH).optional(),
  start: z.number().int().min(0).max(100_000).optional(),
  restart: z.string().max(MAX_FIELD_LENGTH).optional(),
  pos: z.string().max(MAX_FIELD_LENGTH).optional(),
})

export const documentNotesSchema = z.object({
  footnotePr: notePrSchema.optional(),
  endnotePr: notePrSchema.optional(),
})

/**
 * `docProps/core.xml` and `docProps/app.xml`. They go into the `.sdoc` and the user's file, so they
 * are checked on input.
 */
const propertyText = z.string().max(MAX_PROPERTY_LENGTH).optional()

const count = z.number().int().min(0).max(1_000_000_000)

/** The counts saving sends to a DOCX's `app.xml`; see `DocumentStatistics`. */
export const documentStatisticsSchema = z.object({
  pages: count,
  words: count,
  characters: count,
  charactersWithSpaces: count,
  paragraphs: count,
})

export const documentPropertiesSchema = z.object({
  title: propertyText,
  subject: propertyText,
  creator: propertyText,
  keywords: propertyText,
  category: propertyText,
  contentStatus: propertyText,
  language: propertyText,
  lastPrinted: z.string().max(MAX_FIELD_LENGTH).optional(),
  description: z.string().max(MAX_DESCRIPTION_LENGTH).optional(),
  lastModifiedBy: propertyText,
  revision: z.string().max(MAX_FIELD_LENGTH).optional(),
  created: z.string().max(MAX_FIELD_LENGTH).optional(),
  modified: z.string().max(MAX_FIELD_LENGTH).optional(),
  company: propertyText,
  manager: propertyText,
  totalTime: count.optional(),
  removePersonalInformation: z.boolean().optional(),
  custom: z
    .array(
      z.object({
        name: z.string().min(1).max(MAX_NAME_LENGTH),
        type: z.enum(CUSTOM_PROPERTY_TYPES),
        value: z.string().max(MAX_PROPERTY_LENGTH),
      }),
    )
    .max(MAX_CUSTOM_PROPERTIES)
    .readonly()
    .optional(),
})

/**
 * A PDF's information dictionary and XMP, as the sidecar reads them. `application` and `producer`
 * are read only; an encrypted PDF comes without properties.
 */
export const pdfInspectionSchema = z.object({
  version: z.string().max(MAX_FIELD_LENGTH),
  encrypted: z.boolean(),
  signed: z.boolean(),
  properties: documentPropertiesSchema.optional(),
  application: propertyText,
  producer: propertyText,
  /** Incremental updates appended so far; each one leaves the previous revision in the file. */
  updates: z.number().int().nonnegative(),
})

export const privacyCategorySchema = z.enum(PRIVACY_CATEGORIES)

const MAX_PRIVACY_ITEMS = 500

/** What the sidecar finds in a DOCX or XLSX, and the editor in a model of its own. */
export const privacyReportSchema = z.object({
  names: z.array(propertyText.unwrap()).max(MAX_PRIVACY_ITEMS),
  paths: z.array(z.string().max(MAX_PATH_LENGTH)).max(MAX_PRIVACY_ITEMS),
  customProperties: z.array(propertyText.unwrap()).max(MAX_PRIVACY_ITEMS),
  thumbnail: z.boolean(),
  imagesWithMetadata: z.number().int().nonnegative(),
  removePersonalInformation: z.boolean(),
})

/**
 * Lives here, not in the IPC contract, because the same schema serves three places: reading the
 * file where main stores them, the renderer request and the echo back. Three definitions would
 * drift, and the drift would show as a checked menu item the editor ignores.
 *
 * The `default`s let an old installation's file open, which has none of these keys.
 */
export const editorPreferencesSchema = z.object({
  spellcheck: z.boolean().default(true),
  invisibleCharacters: z.boolean().default(false),
  typography: z.boolean().default(true),
  // The `default` is only the parse safety net. On first run the operating system decides, in
  // `src/main/preferences.ts`, which can tell "missing key" from "key saved with this value", a
  // distinction a `default` erases.
  language: z.enum(LANGUAGES).default(Language.Portuguese),
  theme: z.enum([Theme.System, Theme.Light, Theme.Dark]).default(Theme.System),
  readingMode: z.boolean().default(false),
  showToolbar: z.boolean().default(true),
  showStatusBar: z.boolean().default(true),
  zoom: z.number().int().min(MIN_ZOOM).max(MAX_ZOOM).default(100),
  zoomFit: z.boolean().default(false),
  navigationPane: z.boolean().default(false),
  commentsPane: z.boolean().default(true),
  authorName: z.string().max(MAX_NAME_LENGTH).default(''),
})

/**
 * One or more keys, and **only** those sent.
 *
 * Written by hand instead of `editorPreferencesSchema.partial()`, and the contract test exists
 * because of it: `.partial()` makes keys optional but **keeps the `default`s**, so a "show marks"
 * request would come back from the parse with the other keys filled with defaults, and turning
 * spelling off would be undone by the next click on any other key.
 */
export const editorPreferencesPatchSchema = z.object({
  spellcheck: z.boolean().optional(),
  invisibleCharacters: z.boolean().optional(),
  typography: z.boolean().optional(),
  language: z.enum(LANGUAGES).optional(),
  theme: z.enum([Theme.System, Theme.Light, Theme.Dark]).optional(),
  readingMode: z.boolean().optional(),
  showToolbar: z.boolean().optional(),
  showStatusBar: z.boolean().optional(),
  zoom: z.number().int().min(MIN_ZOOM).max(MAX_ZOOM).optional(),
  zoomFit: z.boolean().optional(),
  navigationPane: z.boolean().optional(),
  commentsPane: z.boolean().optional(),
  authorName: z.string().max(MAX_NAME_LENGTH).optional(),
})

/**
 * What Chromium reports about the right-click point. `dictionarySuggestions` feeds menu items, and
 * a long list would leave the screen; Chromium sends five.
 */
export const contextMenuTargetSchema = z.object({
  x: z.number().int().min(0).max(100_000),
  y: z.number().int().min(0).max(100_000),
  editable: z.boolean(),
  misspelledWord: z.string().max(MAX_NAME_LENGTH),
  dictionarySuggestions: z.array(z.string().max(MAX_NAME_LENGTH)).max(10),
  canCut: z.boolean(),
  canCopy: z.boolean(),
  canPaste: z.boolean(),
})

const lineSpacingSchema = z.discriminatedUnion('kind', [
  z.object({ kind: z.literal('multiple'), factor: z.number().min(0).max(100) }),
  z.object({ kind: z.literal('exact'), pt: z.number().min(0).max(2000) }),
  z.object({ kind: z.literal('atLeast'), pt: z.number().min(0).max(2000) }),
])

const styleParagraphSchema = z.object({
  textAlign: z.string().max(MAX_CSS_VALUE_LENGTH).optional(),
  indentMm: z.number().optional(),
  indentRightMm: z.number().optional(),
  firstLineMm: z.number().optional(),
  spaceBefore: z.number().optional(),
  spaceAfter: z.number().optional(),
  lineSpacing: lineSpacingSchema.optional(),
  keepNext: z.boolean().optional(),
  keepLines: z.boolean().optional(),
  widowControl: z.boolean().optional(),
  pageBreakBefore: z.boolean().optional(),
  contextualSpacing: z.boolean().optional(),
  // 9 is "body text": the level Word writes in the `TOC Heading` style, which inherits from
  // `heading 1` and must switch the inherited level off. Refusing it would refuse the stylesheet of
  // every Word document with a table of contents.
  outlineLevel: z.number().int().min(0).max(9).optional(),
  background: z.string().max(MAX_COLOR_LENGTH).optional(),
})

const styleCharacterSchema = z.object({
  fontFamily: z.string().max(MAX_NAME_LENGTH).optional(),
  fontSize: z.string().max(MAX_CSS_VALUE_LENGTH).optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strike: z.boolean().optional(),
  allCaps: z.boolean().optional(),
  smallCaps: z.boolean().optional(),
  verticalAlign: z.string().max(MAX_CSS_VALUE_LENGTH).optional(),
  color: z.string().max(MAX_COLOR_LENGTH).optional(),
  highlight: z.string().max(MAX_COLOR_LENGTH).optional(),
})

/**
 * The three switches have defaults because they derive from the presence of an element in the file:
 * a hand-edited `.sdoc` without them is a style that neither hides nor recommends anything, not an
 * invalid document.
 */
const styleDefinitionSchema = z.object({
  id: z.string().min(1).max(MAX_ID_LENGTH),
  name: z.string().min(1).max(MAX_NAME_LENGTH),
  type: z.enum(['paragraph', 'character']),
  qFormat: z.boolean().default(false),
  hidden: z.boolean().default(false),
  custom: z.boolean().default(false),
  basedOn: z.string().max(MAX_ID_LENGTH).optional(),
  next: z.string().max(MAX_ID_LENGTH).optional(),
  link: z.string().max(MAX_ID_LENGTH).optional(),
  uiPriority: z.number().int().min(0).max(1000).optional(),
  paragraph: styleParagraphSchema.optional(),
  character: styleCharacterSchema.optional(),
})

/**
 * They cross IPC both ways: from the sidecar when opening a `.docx` and back inside the `.sdoc`
 * when saving. Same schema at both points because it is the same data, and a malformed style must
 * not reach the screen.
 *
 * Nothing here is `strict`: a style `w:pPr` has dozens of properties and the reader reads the ones
 * it knows. Refusing the document over a new key would trade an incomplete screen for no screen.
 */
export const styleSheetSchema = z.object({
  defaults: z.object({
    paragraph: styleParagraphSchema.default({}),
    character: styleCharacterSchema.default({}),
    /** The style that applies without `w:pStyle`; `null` when the document marks none. */
    paragraphStyleId: z.string().max(MAX_ID_LENGTH).nullable().default(null),
    characterStyleId: z.string().max(MAX_ID_LENGTH).nullable().default(null),
  }),
  // A net against pathological files, not a design limit: a Word document with a style for every
  // table variant passes 400.
  styles: z
    .record(z.string().max(MAX_ID_LENGTH), styleDefinitionSchema)
    .refine((styles) => Object.keys(styles).length <= 4000, 'estilos demais'),
})

const aiProviderIdSchema = z.string().min(1).max(MAX_ID_LENGTH)
export const aiModelIdSchema = z.string().max(MAX_ID_LENGTH)
const aiBaseUrlSchema = z.url({ protocol: /^https?$/ }).max(MAX_AI_URL_LENGTH)
/** In characters: how much of the document goes along with the selection. */
export const aiContextLengthSchema = z.int().min(0).max(MAX_AI_DOCUMENT_LENGTH)
const aiDocumentPartSchema = z.string().max(MAX_AI_DOCUMENT_LENGTH)
const aiAgentIdSchema = z.enum(ACP_AGENT_IDS)
/** As typed: main splits them, quotes included, and hands them to the agent without a shell. */
export const aiAgentArgsSchema = z.string().max(MAX_AI_AGENT_ARGS_LENGTH)

/** What the renderer sees of a provider: whether it has a key, never the key. */
export const aiProviderSchema = z.object({
  id: aiProviderIdSchema,
  kind: z.enum(AiProviderKind),
  label: z.string().min(1).max(MAX_NAME_LENGTH),
  baseUrl: aiBaseUrlSchema.optional(),
  hasKey: z.boolean(),
  local: z.boolean(),
  consented: z.boolean(),
  defaultModel: aiModelIdSchema.optional(),
  contextLength: aiContextLengthSchema.optional(),
  agentId: aiAgentIdSchema.optional(),
  /** The file name only: the full path of a chosen agent stays in main. */
  commandName: z.string().max(MAX_PATH_LENGTH).optional(),
  args: aiAgentArgsSchema.optional(),
})

export const aiProviderDraftSchema = z.object({
  id: aiProviderIdSchema.optional(),
  kind: z.enum(AiProviderKind),
  label: z.string().trim().min(1).max(MAX_NAME_LENGTH),
  baseUrl: aiBaseUrlSchema.optional(),
  apiKey: z.string().trim().max(MAX_AI_KEY_LENGTH).optional(),
  defaultModel: aiModelIdSchema.optional(),
  contextLength: aiContextLengthSchema.optional(),
  agentId: aiAgentIdSchema.optional(),
  args: aiAgentArgsSchema.optional(),
  makeDefault: z.boolean().default(false),
  revokeConsent: z.boolean().default(false),
})

export const aiPreferencesSchema = z.object({
  /** For the actions that revise the passage: rewrite, proofread, translate. */
  keepFormatting: z.boolean(),
  anthropicFallback: z.boolean(),
  /** In the chat, the document tools: the model proposes edits and comments that the person applies. */
  chatProposals: z.boolean(),
  /** Programs outside the app read the open document and leave revisions and comments, through MCP. */
  mcpServer: z.boolean(),
})

export const aiActionSchema = z.enum(AiAction)

export const aiGenerateRequestSchema = z.object({
  requestId: z.uuid(),
  providerId: aiProviderIdSchema,
  model: aiModelIdSchema,
  action: aiActionSchema,
  targetLanguage: z.enum(LANGUAGES).optional(),
  instruction: z.string().max(MAX_AI_INSTRUCTION_LENGTH).optional(),
  text: z.string().min(1).max(MAX_AI_SELECTION_LENGTH),
  /** The document text before and after the selection, deletions left out. */
  before: aiDocumentPartSchema.default(''),
  after: aiDocumentPartSchema.default(''),
  format: z.enum(AiFormat).default(AiFormat.Plain),
})

const aiChatTurnSchema = z.discriminatedUnion('role', [
  z.object({ role: z.literal(AiRole.User), text: z.string().min(1).max(MAX_AI_CHAT_MESSAGE_LENGTH) }),
  z.object({ role: z.literal(AiRole.Assistant), text: z.string().min(1).max(MAX_AI_RESPONSE_LENGTH) }),
])

/** One text per paragraph, deletions left out. */
const aiParagraphsSchema = z
  .array(aiDocumentPartSchema)
  .max(MAX_AI_PARAGRAPHS)
  .refine(
    (paragraphs) => paragraphs.reduce((sum, text) => sum + text.length, 0) <= MAX_AI_DOCUMENT_LENGTH,
    'document too long',
  )

export const aiChatRequestSchema = z.object({
  requestId: z.uuid(),
  providerId: aiProviderIdSchema,
  model: aiModelIdSchema,
  messages: z
    .array(aiChatTurnSchema)
    .min(1)
    .max(MAX_AI_CHAT_MESSAGES)
    .refine(
      (messages) => messages[messages.length - 1]?.role === AiRole.User,
      'the conversation must end with the person',
    ),
  /** The document text before and after the selection, deletions left out; all of it in `before` without one. */
  before: aiDocumentPartSchema.default(''),
  after: aiDocumentPartSchema.default(''),
  selection: z.string().max(MAX_AI_SELECTION_LENGTH).default(''),
  /** The person asked for proposals in this message; main still checks the preference and the provider. */
  proposals: z.boolean().default(false),
  /** One text per paragraph, deletions left out; sent only when the tools are on. */
  paragraphs: aiParagraphsSchema.default([]),
})

/** Each task brings what it needs: the formula to explain, the description, the rows to fill. */
function missingForTask(request: {
  task: AiSheetTask
  instruction: string
  formula?: string | undefined
  rowCount?: number | undefined
}): string | null {
  if (request.task === AiSheetTask.Explain) return request.formula === undefined ? 'formula' : null
  if (request.instruction.trim() === '') return 'instruction'
  if (request.task === AiSheetTask.Fill && request.rowCount === undefined) return 'rowCount'
  return null
}

export const aiSheetRequestSchema = z
  .object({
    requestId: z.uuid(),
    providerId: aiProviderIdSchema,
    model: aiModelIdSchema,
    task: z.enum(AiSheetTask),
    language: z.enum(LANGUAGES),
    instruction: z.string().max(MAX_AI_INSTRUCTION_LENGTH).default(''),
    formula: z.string().min(1).max(MAX_AI_FORMULA_LENGTH).optional(),
    /** The active sheet only, as text: its name, the headers and the cells that matter. */
    context: z.string().max(MAX_AI_SHEET_CONTEXT_LENGTH),
    rowCount: z.int().min(1).max(MAX_AI_FILL_ROWS).optional(),
  })
  .superRefine((request, context) => {
    const missing = missingForTask(request)
    if (missing !== null)
      context.addIssue({ code: 'custom', path: [missing], message: `${missing} required` })
  })

const aiProposalAnchor = {
  id: z.string().min(1).max(MAX_ID_LENGTH),
  paragraph: z.int().min(1).max(MAX_AI_PARAGRAPHS),
  paragraphText: aiDocumentPartSchema,
  find: z.string().min(1).max(MAX_AI_FIND_LENGTH),
}

export const aiProposalSchema = z.discriminatedUnion('kind', [
  z.object({
    ...aiProposalAnchor,
    kind: z.literal(AiProposalKind.Edit),
    replace: z.string().max(MAX_AI_REPLACE_LENGTH),
    reason: z.string().max(MAX_AI_REASON_LENGTH).optional(),
  }),
  z.object({
    ...aiProposalAnchor,
    kind: z.literal(AiProposalKind.Comment),
    comment: z.string().min(1).max(MAX_AI_COMMENT_LENGTH),
  }),
])

export const aiStreamEventSchema = z.object({
  requestId: z.uuid(),
  delta: z.string().min(1).max(MAX_AI_RESPONSE_LENGTH),
})

const mcpClientNameSchema = z.string().min(1).max(MAX_MCP_CLIENT_NAME_LENGTH)

export const mcpClientsSchema = z.object({ clients: z.array(mcpClientNameSchema).max(MAX_MCP_CLIENTS) })

export const mcpDocumentRequestSchema = z.object({
  requestId: z.uuid(),
  ask: z.enum(McpDocumentAsk),
  /** Epoch milliseconds; past it the window refuses. */
  deadline: z.number().int().positive(),
  /** Only to apply: the proposal goes in at once, by this author. */
  proposal: aiProposalSchema.optional(),
  author: z.string().min(1).max(MAX_NAME_LENGTH).optional(),
})

export const mcpDocumentReplySchema = z.discriminatedUnion('ask', [
  z.object({ requestId: z.uuid(), ask: z.literal(McpDocumentAsk.Snapshot), paragraphs: aiParagraphsSchema }),
  z.object({ requestId: z.uuid(), ask: z.literal(McpDocumentAsk.Apply), applied: z.boolean() }),
  z.object({
    requestId: z.uuid(),
    ask: z.literal(MCP_DOCUMENT_REFUSED),
    refusal: z.enum(McpDocumentRefusal),
  }),
])
