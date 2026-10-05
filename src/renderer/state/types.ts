import type { SerializedError } from '@shared/errors.js'
import type { DocumentKind, DraftSummary, LossInventory, RecentFile } from '@shared/types.js'
import type {
  DocumentComment,
  DocumentModel,
  DocumentNode,
  DocumentNotes,
  DocumentProperties,
  PageSetup,
  SectionSetup,
} from '@services/document/model.js'
import type { PagedDocument } from '@services/document/print-pages.js'
import type { StyleSheet } from '@services/document/styles.js'
import type { Sheet, WorkbookModel } from '@services/spreadsheet/model.js'
import type { StructuralChange } from '@services/spreadsheet/structure.js'

export interface OpenFile {
  /** `null` while the file was never saved. */
  readonly path: string | null
  /**
   * The Word template the new document came from: the origin of the first save, never a
   * destination.
   */
  readonly origin?: string
  readonly name: string
  readonly kind: DocumentKind
}

/** The app edits one file at a time: `workbook` says which of the two applies. */
export interface LoadedFile {
  readonly file: OpenFile
  readonly model: DocumentModel
  readonly workbook: WorkbookModel | null
}

/** The HTML comes from the editor itself, so the PDF matches the screen. */
export interface DocumentSource {
  readonly readDoc: () => DocumentNode
  readonly readHtml: () => string
  /** Paper comes from here, not from a second pagination. */
  readonly readPages: () => PagedDocument
}

export interface WorkspaceState {
  file: OpenFile | null
  page: PageSetup
  /** Does not follow typing: the live content lives in the editor. */
  initialDoc: DocumentNode
  styles: StyleSheet
  /** Blocks came flattened from an old draft; see `DocumentModel.flattened`. */
  flattened: boolean
  /** The draft predates references; see `DocumentModel.beforeReferences`. */
  beforeReferences: boolean
  /** Sections before the last; see `DocumentModel.sections`. Empty means a single section. */
  sections: readonly SectionSetup[]
  /** The draft predates sections; see `DocumentModel.beforeSections`. */
  beforeSections: boolean
  /** File bookmarks outside the nodes; see `DocumentModel.outsideBookmarks`. */
  outsideBookmarks: readonly string[]
  /**
   * Only gains entries: the text says which count (`resolveComments`), and undo removes and
   * restores.
   */
  comments: readonly DocumentComment[]
  /** As conversas ancoradas fora do corpo — ver `commentsOutsideOf`. */
  commentsOutside: readonly string[]
  /** The just-inserted comment whose box the pane opens. */
  commentDraft: string | null
  /** The draft predates comments; see `DocumentModel.beforeComments`. */
  beforeComments: boolean
  /** The file's `w:trackRevisions`; see `DocumentModel.trackChanges`. */
  trackChanges: boolean | undefined
  /** The draft predates revisions; see `DocumentModel.beforeRevisions`. */
  beforeRevisions: boolean
  /** See `DocumentModel.notes`. */
  notes: DocumentNotes | undefined
  /** The draft predates notes; see `DocumentModel.beforeNotes`. */
  beforeNotes: boolean
  /** The draft predates equations; see `DocumentModel.beforeMath`. */
  beforeMath: boolean
  properties: DocumentProperties | undefined
  workbook: WorkbookModel | null
  /** Changes on every new/open: the editor is remounted, with no leftover state. */
  generation: number
  isDirty: boolean
  stats: { characters: number; words: number }
  pageCount: number
  recents: readonly RecentFile[]
  error: SerializedError | null
  /** Not a failure: the file opened, and this says what will not show. */
  notice: LossInventory | null
  /** What the last save lost: `notice` warns what **will** be lost, and this what **was** lost. */
  savedLoss: readonly string[] | null
  /** While it awaits a decision, autosave does not write over it. */
  pendingDraft: DraftSummary | null
  /** Turned on by itself when there is **structural** loss; the user unlocks with one click. */
  readOnly: boolean
  /** Retrying every eight seconds would fill the screen with warnings. */
  autosaveBroken: boolean
  busy: boolean

  registerDocumentSource: (source: DocumentSource | null) => void
  markDirty: () => void
  setStats: (stats: { characters: number; words: number }) => void
  setPageCount: (pages: number) => void
  setPage: (page: PageSetup) => void
  /** The sections' setup or the whole list. Like `setPage`, it marks the document. */
  setSections: (sections: readonly SectionSetup[]) => void
  /** Outside the editor undo, like styles; marks the document. */
  setComments: (comments: readonly DocumentComment[]) => void
  setCommentDraft: (cid: string | null) => void
  /** Always a new sheet: the new reference is what regenerates the style CSS. */
  setStyles: (styles: StyleSheet) => void
  /** It belongs to the document (`w:trackRevisions`), not the person: marks the document. */
  toggleTrackChanges: () => void
  /** Outside the editor undo, like styles; marks the document. */
  setProperties: (properties: DocumentProperties) => void
  dismissError: () => void
  dismissNotice: () => void
  showError: (error: SerializedError) => void
  refreshRecents: () => Promise<void>

  newDocument: () => Promise<void>
  newSpreadsheet: () => Promise<void>
  updateSheet: (sheet: Sheet) => void
  changeStructure: (change: StructuralChange) => void
  selectSheet: (index: number) => void
  addSheet: () => void
  renameSheet: (index: number, name: string) => void
  removeSheet: (index: number) => void
  openViaDialog: () => Promise<void>
  openRecent: (path: string) => Promise<void>
  templateGallery: boolean
  setTemplateGallery: (open: boolean) => void
  /** `null` browses for a `.dotx`. Returns whether the new document reached the screen. */
  newFromTemplate: (
    template: { readonly source: 'builtin' | 'user'; readonly id: string } | null,
  ) => Promise<boolean>
  save: () => Promise<boolean>
  saveAs: () => Promise<boolean>
  closeFile: () => Promise<void>
  clearRecents: () => Promise<void>

  /** By timer, not by key press. */
  autosave: () => Promise<void>
  checkRecovery: () => Promise<void>
  recoverDraft: () => Promise<void>
  dismissDraft: () => Promise<void>
  allowEditing: () => void

  exportPdf: () => Promise<boolean>
  /** To a new file: the document stays at its path, with the modified state it had. */
  exportDocument: (format: 'html' | 'markdown' | 'odt') => Promise<boolean>
  print: () => Promise<boolean>
  printPreview: () => Promise<void>
}
