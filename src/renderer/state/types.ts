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
  /** `null` enquanto o arquivo nunca foi gravado. */
  readonly path: string | null
  /** O modelo do Word de que o documento novo saiu: origem da primeira gravação, nunca destino. */
  readonly origin?: string
  readonly name: string
  readonly kind: DocumentKind
}

/** O aplicativo edita um arquivo por vez: `workbook` diz qual dos dois vale. */
export interface LoadedFile {
  readonly file: OpenFile
  readonly model: DocumentModel
  readonly workbook: WorkbookModel | null
}

/** O HTML vem do próprio editor, para o PDF sair igual à tela. */
export interface DocumentSource {
  readonly readDoc: () => DocumentNode
  readonly readHtml: () => string
  /** O papel sai daqui, e não de uma segunda paginação. */
  readonly readPages: () => PagedDocument
}

export interface WorkspaceState {
  file: OpenFile | null
  page: PageSetup
  /** Não acompanha a digitação: o conteúdo ao vivo mora no editor. */
  initialDoc: DocumentNode
  styles: StyleSheet
  /** Os blocos vieram achatados de um rascunho antigo — ver `DocumentModel.flattened`. */
  flattened: boolean
  /** O rascunho é de antes das referências — ver `DocumentModel.beforeReferences`. */
  beforeReferences: boolean
  /** As seções antes da última — ver `DocumentModel.sections`. Vazio é uma seção só. */
  sections: readonly SectionSetup[]
  /** O rascunho é de antes das seções — ver `DocumentModel.beforeSections`. */
  beforeSections: boolean
  /** Marcadores do arquivo fora dos nós — ver `DocumentModel.outsideBookmarks`. */
  outsideBookmarks: readonly string[]
  /** Só ganha entradas: o texto diz quais valem (`resolveComments`), e o desfazer tira e devolve. */
  comments: readonly DocumentComment[]
  /** As conversas ancoradas fora do corpo — ver `commentsOutsideOf`. */
  commentsOutside: readonly string[]
  /** O comentário recém-inserido cuja caixa o painel abre. */
  commentDraft: string | null
  /** O rascunho é de antes dos comentários — ver `DocumentModel.beforeComments`. */
  beforeComments: boolean
  /** O `w:trackRevisions` do arquivo — ver `DocumentModel.trackChanges`. */
  trackChanges: boolean | undefined
  /** O rascunho é de antes das revisões — ver `DocumentModel.beforeRevisions`. */
  beforeRevisions: boolean
  /** A numeração das notas do documento — ver `DocumentModel.notes`. */
  notes: DocumentNotes | undefined
  /** O rascunho é de antes das notas — ver `DocumentModel.beforeNotes`. */
  beforeNotes: boolean
  /** O rascunho é de antes das equações — ver `DocumentModel.beforeMath`. */
  beforeMath: boolean
  properties: DocumentProperties | undefined
  workbook: WorkbookModel | null
  /** Muda a cada novo/abrir: o editor é remontado, sem estado residual. */
  generation: number
  isDirty: boolean
  stats: { characters: number; words: number }
  pageCount: number
  recents: readonly RecentFile[]
  error: SerializedError | null
  /** Não é falha: o arquivo abriu, e isto diz o que não vai aparecer. */
  notice: LossInventory | null
  /** O que a última gravação perdeu: o `notice` avisa o que **vai** se perder, e este o que **se perdeu**. */
  savedLoss: readonly string[] | null
  /** Enquanto ele espera decisão, o autosave não escreve por cima dele. */
  pendingDraft: DraftSummary | null
  /** Ligado sozinho quando há perda **estrutural**; a pessoa libera num clique. */
  readOnly: boolean
  /** Insistir a cada oito segundos encheria a tela de avisos. */
  autosaveBroken: boolean
  busy: boolean

  registerDocumentSource: (source: DocumentSource | null) => void
  markDirty: () => void
  setStats: (stats: { characters: number; words: number }) => void
  setPageCount: (pages: number) => void
  setPage: (page: PageSetup) => void
  /** A configuração das seções ou a lista inteira. Como `setPage`, marca o documento. */
  setSections: (sections: readonly SectionSetup[]) => void
  /** Fora do desfazer do editor, como os estilos; marca o documento. */
  setComments: (comments: readonly DocumentComment[]) => void
  setCommentDraft: (cid: string | null) => void
  /** Sempre uma folha nova: é a referência nova que regera o CSS dos estilos. */
  setStyles: (styles: StyleSheet) => void
  /** É do documento (`w:trackRevisions`), e não da pessoa: marca o documento. */
  toggleTrackChanges: () => void
  /** Fora do desfazer do editor, como os estilos; marca o documento. */
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
  /** `null` procura um `.dotx`. Devolve se o documento novo chegou à tela. */
  newFromTemplate: (
    template: { readonly source: 'builtin' | 'user'; readonly id: string } | null,
  ) => Promise<boolean>
  save: () => Promise<boolean>
  saveAs: () => Promise<boolean>
  closeFile: () => Promise<void>
  clearRecents: () => Promise<void>

  /** Por relógio, e não por tecla. */
  autosave: () => Promise<void>
  checkRecovery: () => Promise<void>
  recoverDraft: () => Promise<void>
  dismissDraft: () => Promise<void>
  allowEditing: () => void

  exportPdf: () => Promise<boolean>
  /** Num arquivo novo: o documento continua no caminho dele, com o estado de alterado que tinha. */
  exportDocument: (format: 'html' | 'markdown' | 'odt') => Promise<boolean>
  print: () => Promise<boolean>
  printPreview: () => Promise<void>
}
