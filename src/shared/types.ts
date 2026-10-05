import { Language } from './i18n/language.js'

/** Drives the icon, dialog filters and editor. */
export const DocumentKind = {
  Document: 'document',
  Spreadsheet: 'spreadsheet',
} as const

export type DocumentKind = (typeof DocumentKind)[keyof typeof DocumentKind]

/**
 * `invisible` stays in the file after saving but does not show on screen; `lost` disappears on
 * save. Kept apart because people learn to ignore a generic warning.
 */
export interface LossInventory {
  // Mutable collections, as in `DocumentNode`: this type must be assignable to what zod infers in
  // the IPC contract, and a `readonly` array is not assignable to a plain one.
  readonly invisible: string[]
  readonly lost: string[]
  /**
   * Subset of `invisible`: what disappears if the block anchoring it is edited. Decides whether the
   * file opens read-only.
   */
  readonly structural: string[]
}

/**
 * `content` is always text. A `.docx` arrives **already converted** to the internal format; the
 * original bytes stay in main, which writes surgically.
 */
export interface LoadedFile {
  readonly path: string
  readonly name: string
  readonly kind: DocumentKind
  readonly content: string
  /** Only for Office files. */
  readonly inventory?: LossInventory
  /** A Word template opens as a new document; see `ipc.ts`. */
  readonly template?: boolean
}

export interface TemplateEntry {
  readonly source: 'builtin' | 'user'
  /** The file name for builtin templates; the path for the user's. */
  readonly id: string
  readonly name: string
  readonly description: string
}

/**
 * Without the content, which can be tens of megabytes: the prompt only needs the origin and date.
 */
export interface DraftSummary {
  readonly path: string | null
  readonly name: string
  readonly kind: DocumentKind
  readonly savedAt: number
}

export interface RecentFile {
  readonly path: string
  readonly name: string
  readonly kind: DocumentKind
  readonly openedAt: number
}

export const MenuCommand = {
  NewDocument: 'new-document',
  NewSpreadsheet: 'new-spreadsheet',
  NewFromTemplate: 'new-from-template',
  Open: 'open',
  OpenRecent: 'open-recent',
  ClearRecent: 'clear-recent',
  Save: 'save',
  SaveAs: 'save-as',
  CloseFile: 'close-file',
  FindReplace: 'find-replace',
  ExportPdf: 'export-pdf',
  ExportHtml: 'export-html',
  ExportMarkdown: 'export-markdown',
  ExportOdt: 'export-odt',
  Print: 'print',
  PrintPreview: 'print-preview',
  PageSetup: 'page-setup',
  ParagraphSetup: 'paragraph-setup',
  InsertPageBreak: 'insert-page-break',
  /** A section break at the cursor, at the start of the new section. */
  InsertSectionNextPage: 'insert-section-next-page',
  InsertSectionContinuous: 'insert-section-continuous',
  InsertSectionEvenPage: 'insert-section-even-page',
  InsertSectionOddPage: 'insert-section-odd-page',
  /** The section above takes over the one below. */
  DeleteSectionBreak: 'delete-section-break',
  InsertColumnBreak: 'insert-column-break',
  FormatColumns: 'format-columns',
  /** The editor's zoom, not Chromium's, which would scale the UI. */
  ZoomIn: 'zoom-in',
  ZoomOut: 'zoom-out',
  ZoomReset: 'zoom-reset',
  ZoomFitWidth: 'zoom-fit-width',
  PasteWithoutFormat: 'paste-without-format',
  WordCount: 'word-count',
  DocumentProperties: 'document-properties',
  SpecialCharacter: 'special-character',
  InsertEquation: 'insert-equation',
  InsertDisplayEquation: 'insert-display-equation',
  /** Double click and Enter end up here. */
  EditEquation: 'edit-equation',
  ImageProperties: 'image-properties',
  InsertBookmark: 'insert-bookmark',
  InsertComment: 'insert-comment',
  InsertFootnote: 'insert-footnote',
  InsertEndnote: 'insert-endnote',
  NextComment: 'next-comment',
  PreviousComment: 'previous-comment',
  AcceptChange: 'accept-change',
  RejectChange: 'reject-change',
  AcceptAllChanges: 'accept-all-changes',
  RejectAllChanges: 'reject-all-changes',
  NextChange: 'next-change',
  PreviousChange: 'previous-change',
  ToggleTrackChanges: 'toggle-track-changes',
  ShowAllMarkup: 'show-all-markup',
  ShowSimpleMarkup: 'show-simple-markup',
  ShowNoMarkup: 'show-no-markup',
  ShowOriginal: 'show-original',
  AuthorName: 'author-name',
  InsertTableOfContents: 'insert-table-of-contents',
  UpdateTableOfContents: 'update-table-of-contents',
  UpdateFields: 'update-fields',
  InsertCaption: 'insert-caption',
  InsertCrossReference: 'insert-cross-reference',
  // The values match `TableAction` (see `table-actions.ts`), and `App` forwards them to the editor
  // by them.
  TableInsert: 'table-insert',
  TableRowBefore: 'table-row-before',
  TableRowAfter: 'table-row-after',
  TableDeleteRow: 'table-delete-row',
  TableColumnBefore: 'table-column-before',
  TableColumnAfter: 'table-column-after',
  TableDeleteColumn: 'table-delete-column',
  TableMergeCells: 'table-merge-cells',
  TableSplitCell: 'table-split-cell',
  TableHeaderRow: 'table-header-row',
  TableDelete: 'table-delete',
  TableProperties: 'table-properties',
  /** When the user chooses "Save" in the exit prompt. */
  SaveAndExit: 'save-and-exit',
} as const

export type MenuCommand = (typeof MenuCommand)[keyof typeof MenuCommand]

/** Belongs to the window, as in Word: nothing of this is saved, and printing follows the window. */
export const RevisionView = {
  All: 'all',
  Simple: 'simple',
  None: 'none',
  Original: 'original',
} as const

export type RevisionView = (typeof RevisionView)[keyof typeof RevisionView]

export const DiscardChoice = {
  Save: 'save',
  Discard: 'discard',
  Cancel: 'cancel',
} as const

export type DiscardChoice = (typeof DiscardChoice)[keyof typeof DiscardChoice]

export const PlainTextChoice = {
  /** Accepting the formatting loss. */
  KeepPlain: 'keep-plain',
  SaveAsDocument: 'save-as-document',
  Cancel: 'cancel',
} as const

export type PlainTextChoice = (typeof PlainTextChoice)[keyof typeof PlainTextChoice]

/** Main resolves `system`, since it sees Chromium's `nativeTheme`. */
export const Theme = {
  System: 'system',
  Light: 'light',
  Dark: 'dark',
} as const

export type Theme = (typeof Theme)[keyof typeof Theme]

export type ResolvedTheme = 'light' | 'dark'

/**
 * They live in main: spelling is a `session` setting, the language builds the native menu and the
 * theme needs `nativeTheme`.
 */
export interface EditorPreferences {
  readonly spellcheck: boolean
  /** ¶, space, tab and line break. */
  readonly invisibleCharacters: boolean
  /** Curly quotes, dashes, ellipsis. */
  readonly typography: boolean
  /** Changes neither formulas, which accept both languages, nor the spellchecker dictionary. */
  readonly language: Language
  readonly theme: Theme
  /**
   * A preference, not session state, because the native menu shows the check mark and only knows
   * what is here.
   */
  readonly readingMode: boolean
  readonly showToolbar: boolean
  readonly showStatusBar: boolean
  /** Percent (50–200). Does not change pagination. */
  readonly zoom: number
  /** The zoom follows the window, and `zoom` stays as it was. */
  readonly zoomFit: boolean
  /**
   * A preference for the same reason as `readingMode`; in Word the pane also stays open in the next
   * document.
   */
  readonly navigationPane: boolean
  /** Hidden comments stay in the document and the file. */
  readonly commentsPane: boolean
  /**
   * Signs new comments. Empty in the file, main uses the system user; see `load()` in
   * `src/main/preferences.ts`.
   */
  readonly authorName: string
}

/**
 * `Partial` is not enough with `exactOptionalPropertyTypes`: what zod infers from a partial schema
 * allows the key with `undefined`, and that is the value that crosses IPC.
 */
export type EditorPreferencesPatch = {
  readonly [K in keyof EditorPreferences]?: EditorPreferences[K] | undefined
}

/**
 * `language` is the last resort: on first run, `load()` in `src/main/preferences.ts` uses the
 * system language.
 */
export const DEFAULT_EDITOR_PREFERENCES: EditorPreferences = {
  spellcheck: true,
  invisibleCharacters: false,
  typography: true,
  language: Language.Portuguese,
  theme: Theme.System,
  readingMode: false,
  showToolbar: true,
  showStatusBar: true,
  zoom: 100,
  zoomFit: false,
  navigationPane: false,
  commentsPane: true,
  authorName: '',
}

/** Only `webContents` can do these. */
export const EditCommand = {
  Cut: 'cut',
  Copy: 'copy',
  Paste: 'paste',
} as const

export type EditCommand = (typeof EditCommand)[keyof typeof EditCommand]

/**
 * From the `webContents` `context-menu` event, the only one that carries what Chromium's
 * spellchecker found and suggests.
 */
export interface ContextMenuTarget {
  /** Window pixels. */
  readonly x: number
  readonly y: number
  /** Outside an editable target, paste makes no sense. */
  readonly editable: boolean
  /** Empty when the click did not land on a misspelled word. */
  readonly misspelledWord: string
  readonly dictionarySuggestions: string[]
  readonly canCut: boolean
  readonly canCopy: boolean
  readonly canPaste: boolean
}

/**
 * `session` imitates "ignore", which Chromium lacks: added to the dictionary and removed at the end
 * of the session.
 */
export const DictionaryScope = {
  Permanent: 'permanent',
  Session: 'session',
} as const

export type DictionaryScope = (typeof DictionaryScope)[keyof typeof DictionaryScope]
