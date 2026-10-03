import { Language } from './i18n/language.js'

/** Guia ícone, filtros de diálogo e editor. */
export const DocumentKind = {
  Document: 'document',
  Spreadsheet: 'spreadsheet',
} as const

export type DocumentKind = (typeof DocumentKind)[keyof typeof DocumentKind]

/**
 * `invisible` continua no arquivo depois de salvar, mas não aparece na tela;
 * `lost` some ao salvar. Separados porque um aviso genérico se aprende a ignorar.
 */
export interface LossInventory {
  // Coleções mutáveis, como em `DocumentNode`: este tipo precisa ser atribuível
  // ao que o zod infere no contrato de IPC, e um array `readonly` não é
  // atribuível a um comum. As propriedades continuam `readonly`.
  readonly invisible: string[]
  readonly lost: string[]
  /**
   * Subconjunto de `invisible`: o que some se o bloco que o ancora for editado.
   * Decide se o arquivo abre em somente leitura.
   */
  readonly structural: string[]
}

/**
 * `content` é sempre texto. Um `.docx` chega **já convertido** para o formato
 * interno; os bytes originais ficam no main, que grava cirurgicamente.
 */
export interface LoadedFile {
  readonly path: string
  readonly name: string
  readonly kind: DocumentKind
  readonly content: string
  /** Presente só quando o arquivo veio de um formato do Office. */
  readonly inventory?: LossInventory
  /** O arquivo é um modelo do Word: abre como documento novo — ver `ipc.ts`. */
  readonly template?: boolean
}

export interface TemplateEntry {
  readonly source: 'builtin' | 'user'
  /** O nome do arquivo no embutido; o caminho no do usuário. */
  readonly id: string
  readonly name: string
  readonly description: string
}

/** Sem o conteúdo, que pode ter dezenas de megabytes: o aviso só precisa da origem e da data. */
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

/** Comandos que o menu nativo despacha para o renderer. */
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
  /** Quebra de seção no cursor, pelo começo da seção nova. */
  InsertSectionNextPage: 'insert-section-next-page',
  InsertSectionContinuous: 'insert-section-continuous',
  InsertSectionEvenPage: 'insert-section-even-page',
  InsertSectionOddPage: 'insert-section-odd-page',
  /** Exclui a quebra que fecha a seção do cursor: a seção de cima assume a de baixo. */
  DeleteSectionBreak: 'delete-section-break',
  InsertColumnBreak: 'insert-column-break',
  FormatColumns: 'format-columns',
  /** Zoom da folha: o do editor, e não o do Chromium, que aumentaria a interface. */
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
  /** Abre a equação selecionada no editor — o clique duplo e o Enter chegam aqui. */
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
  /** `w:trackRevisions` do documento. */
  ToggleTrackChanges: 'toggle-track-changes',
  ShowAllMarkup: 'show-all-markup',
  ShowSimpleMarkup: 'show-simple-markup',
  ShowNoMarkup: 'show-no-markup',
  ShowOriginal: 'show-original',
  AuthorName: 'author-name',
  InsertTableOfContents: 'insert-table-of-contents',
  UpdateTableOfContents: 'update-table-of-contents',
  UpdateFields: 'update-fields',
  /** Legenda com número (`SEQ`): Figura 1, Tabela 1… */
  InsertCaption: 'insert-caption',
  /** Referência cruzada a título, marcador ou legenda (`REF`/`PAGEREF`). */
  InsertCrossReference: 'insert-cross-reference',
  // O menu "Tabela". Os valores são os mesmos de `TableAction` (ver
  // `table-actions.ts`), e é por eles que o `App` os repassa ao editor.
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
  /** Emitido quando o usuário escolhe "Salvar" no aviso de saída. */
  SaveAndExit: 'save-and-exit',
} as const

export type MenuCommand = (typeof MenuCommand)[keyof typeof MenuCommand]

/** É da janela, como no Word: nada disso é gravado, e a impressão segue a janela. */
export const RevisionView = {
  All: 'all',
  Simple: 'simple',
  None: 'none',
  Original: 'original',
} as const

export type RevisionView = (typeof RevisionView)[keyof typeof RevisionView]

/** Resposta do aviso de alterações não salvas. */
export const DiscardChoice = {
  Save: 'save',
  Discard: 'discard',
  Cancel: 'cancel',
} as const

export type DiscardChoice = (typeof DiscardChoice)[keyof typeof DiscardChoice]

/** Resposta ao aviso de que `.txt` não guarda formatação. */
export const PlainTextChoice = {
  /** Salvar assim mesmo, aceitando a perda de formatação. */
  KeepPlain: 'keep-plain',
  /** Salvar como documento, preservando tudo. */
  SaveAsDocument: 'save-as-document',
  Cancel: 'cancel',
} as const

export type PlainTextChoice = (typeof PlainTextChoice)[keyof typeof PlainTextChoice]

/** Quem resolve `system` é o main, que enxerga o `nativeTheme` do Chromium. */
export const Theme = {
  System: 'system',
  Light: 'light',
  Dark: 'dark',
} as const

export type Theme = (typeof Theme)[keyof typeof Theme]

/** O tema depois de `system` virar um dos dois de verdade. */
export type ResolvedTheme = 'light' | 'dark'

/**
 * Moram no main: a ortografia é configuração de `session`, o idioma monta o menu
 * nativo e o tema precisa do `nativeTheme`.
 */
export interface EditorPreferences {
  readonly spellcheck: boolean
  /** Marcas de formatação: ¶, espaço, tabulação e quebra de linha. */
  readonly invisibleCharacters: boolean
  /** Autocorreção tipográfica: aspas curvas, travessão, reticências. */
  readonly typography: boolean
  /** Não muda as fórmulas, que aceitam os dois idiomas, nem o dicionário do corretor. */
  readonly language: Language
  readonly theme: Theme
  /** Preferência, e não estado da sessão, porque o menu nativo mostra a marca e só sabe o que está aqui. */
  readonly readingMode: boolean
  readonly showToolbar: boolean
  readonly showStatusBar: boolean
  /** Zoom da folha na tela, em porcento (50–200). Não muda a paginação. */
  readonly zoom: number
  /** Ajustar à largura: o zoom acompanha a janela, e `zoom` fica como estava. */
  readonly zoomFit: boolean
  /** Preferência pelo mesmo motivo de `readingMode`; no Word o painel também fica aberto no documento seguinte. */
  readonly navigationPane: boolean
  /** Escondido, os comentários continuam no documento e no arquivo. */
  readonly commentsPane: boolean
  /**
   * O nome que assina os comentários novos. Vazio no arquivo, o main põe o
   * usuário do sistema — ver `load()` em `src/main/preferences.ts`.
   */
  readonly authorName: string
}

/**
 * `Partial` não basta com `exactOptionalPropertyTypes`: o que o zod infere de um
 * schema parcial admite a chave presente com `undefined`, e é esse valor que
 * atravessa o IPC.
 */
export type EditorPreferencesPatch = {
  readonly [K in keyof EditorPreferences]?: EditorPreferences[K] | undefined
}

/** `language` só vale em último caso: na primeira execução, `load()` em `src/main/preferences.ts` usa o idioma do sistema. */
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

/** Operações de área de transferência que só o `webContents` sabe fazer. */
export const EditCommand = {
  Cut: 'cut',
  Copy: 'copy',
  Paste: 'paste',
} as const

export type EditCommand = (typeof EditCommand)[keyof typeof EditCommand]

/** Vem do evento `context-menu` do `webContents`, o único que traz o que o corretor do Chromium achou e sugere. */
export interface ContextMenuTarget {
  /** Onde clicou, em pixels da janela. */
  readonly x: number
  readonly y: number
  /** O clique caiu em algo editável — fora disso, colar não faz sentido. */
  readonly editable: boolean
  /** Vazio quando o clique não caiu sobre palavra marcada como errada. */
  readonly misspelledWord: string
  readonly dictionarySuggestions: string[]
  readonly canCut: boolean
  readonly canCopy: boolean
  readonly canPaste: boolean
}

/** `session` imita o "ignorar", que o Chromium não tem: entra no dicionário e sai no fim da sessão. */
export const DictionaryScope = {
  Permanent: 'permanent',
  Session: 'session',
} as const

export type DictionaryScope = (typeof DictionaryScope)[keyof typeof DictionaryScope]
