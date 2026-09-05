import type { SerializedError } from '@shared/errors.js'
import type { DocumentKind, DraftSummary, LossInventory, RecentFile } from '@shared/types.js'
import type { DocumentModel, DocumentNode, PageSetup } from '@services/document/model.js'
import type { PagedDocument } from '@services/document/print-pages.js'
import type { Sheet, WorkbookModel } from '@services/spreadsheet/model.js'
import type { StructuralChange } from '@services/spreadsheet/structure.js'

/** O arquivo aberto: o que a barra de título e a gravação precisam saber dele. */
export interface OpenFile {
  /** `null` enquanto o arquivo nunca foi gravado. */
  readonly path: string | null
  readonly name: string
  readonly kind: DocumentKind
}

/**
 * Um arquivo já interpretado, pronto para ir à tela.
 *
 * Os dois campos são excludentes na prática — o aplicativo edita um arquivo por
 * vez —, e é `workbook` quem diz qual dos dois está valendo.
 */
export interface LoadedFile {
  readonly file: OpenFile
  readonly model: DocumentModel
  readonly workbook: WorkbookModel | null
}

/**
 * A porta do editor para quem precisa do conteúdo que está na tela.
 *
 * O HTML vem do próprio editor, e não de uma segunda renderização a partir do
 * modelo — é o que garante que o PDF saia igual ao que está na tela.
 */
export interface DocumentSource {
  readonly readDoc: () => DocumentNode
  readonly readHtml: () => string
  /**
   * O documento recortado nas páginas que a tela está mostrando.
   *
   * O papel sai daqui, e não de uma segunda paginação: era a divergência que o
   * §6.3 do plano registrava como risco residual — dois paginadores, um em
   * JavaScript e outro no Chromium, sem nada forçando a sincronia.
   */
  readonly readPages: () => PagedDocument
}

export interface WorkspaceState {
  file: OpenFile | null
  page: PageSetup
  /**
   * Documento com que o editor é inicializado. Não acompanha a digitação — o
   * conteúdo ao vivo mora no editor, e só é lido na hora de salvar.
   */
  initialDoc: DocumentNode
  /**
   * A planilha aberta, quando o que está em edição é uma planilha.
   *
   * Fica ao lado do documento em vez de substituí-lo porque o aplicativo edita
   * um arquivo por vez: qual dos dois está preenchido é o que diz o tipo.
   */
  workbook: WorkbookModel | null
  /** Muda a cada novo/abrir; o editor é remontado, evitando estado residual. */
  generation: number
  isDirty: boolean
  /** Contagens exibidas na barra de status; alimentadas pelo editor. */
  stats: { characters: number; words: number }
  /**
   * Páginas estimadas, medidas na tela pelo editor. É **estimativa**: a
   * paginação de verdade só acontece na exportação, com regras de viúvas e
   * órfãs que o editor não aplica. Por isso a barra mostra "≈".
   */
  estimatedPages: number
  recents: readonly RecentFile[]
  error: SerializedError | null
  /**
   * O que o documento aberto tem e o editor não dá conta. Fica separado do
   * `error` porque não é falha: o arquivo abriu, e isto é informação sobre o
   * que você vai e o que você não vai ver.
   */
  notice: LossInventory | null
  /**
   * Rascunho de uma sessão que não terminou bem, esperando decisão.
   *
   * Enquanto ele está aqui o autosave não escreve: gravar por cima do rascunho
   * antes de o usuário decidir apagaria justamente o trabalho que ele existe
   * para devolver.
   */
  pendingDraft: DraftSummary | null
  /**
   * O arquivo abriu travado contra edição.
   *
   * Ligado sozinho quando o inventário traz perda **estrutural** — comentário,
   * revisão, nota, campo calculado. Não é um cadeado: é um padrão, e o usuário
   * libera num clique. Travar por perda de aparência travaria o uso do dia a
   * dia, e aí ele aprenderia a liberar sem ler.
   */
  readOnly: boolean
  /**
   * O autosave falhou e parou de tentar.
   *
   * Insistir a cada oito segundos numa gravação que não vai dar certo encheria a
   * tela de avisos; ficar tentando em silêncio deixaria o usuário confiando numa
   * rede de proteção que não existe mais.
   */
  autosaveBroken: boolean
  busy: boolean

  /**
   * O editor se registra aqui para que salvar e imprimir consigam ler o
   * conteúdo atual.
   */
  registerDocumentSource: (source: DocumentSource | null) => void
  markDirty: () => void
  setStats: (stats: { characters: number; words: number }) => void
  setEstimatedPages: (pages: number) => void
  setPage: (page: PageSetup) => void
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
  save: () => Promise<boolean>
  saveAs: () => Promise<boolean>
  closeFile: () => Promise<void>
  clearRecents: () => Promise<void>

  /** Guarda o que está na tela como rascunho. Chamado por relógio, não por tecla. */
  autosave: () => Promise<void>
  checkRecovery: () => Promise<void>
  recoverDraft: () => Promise<void>
  dismissDraft: () => Promise<void>
  /** Libera a edição de um arquivo aberto em somente leitura. */
  allowEditing: () => void

  exportPdf: () => Promise<boolean>
  print: () => Promise<boolean>
  printPreview: () => Promise<void>
}
