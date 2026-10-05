import { z } from 'zod'
import type { DocumentNode } from '@services/document/model.js'
import { MAX_ZOOM, MIN_ZOOM } from '@services/document/zoom.js'
import { Language, LANGUAGES } from './i18n/language.js'
import { Theme } from './types.js'
import {
  MAX_COLOR_LENGTH,
  MAX_CSS_VALUE_LENGTH,
  MAX_DESCRIPTION_LENGTH,
  MAX_FIELD_LENGTH,
  MAX_ID_LENGTH,
  MAX_IMAGE_SIDE_PX,
  MAX_NAME_LENGTH,
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

export const documentPropertiesSchema = z.object({
  title: propertyText,
  subject: propertyText,
  creator: propertyText,
  keywords: propertyText,
  category: propertyText,
  description: z.string().max(MAX_DESCRIPTION_LENGTH).optional(),
  lastModifiedBy: propertyText,
  revision: z.string().max(MAX_FIELD_LENGTH).optional(),
  created: z.string().max(MAX_FIELD_LENGTH).optional(),
  modified: z.string().max(MAX_FIELD_LENGTH).optional(),
  company: propertyText,
  manager: propertyText,
  totalTime: z.number().int().min(0).max(1_000_000_000).optional(),
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
