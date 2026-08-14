/** O que está sendo editado. Guia ícone, filtros de diálogo e editor. */
export const DocumentKind = {
  Document: 'document',
  Spreadsheet: 'spreadsheet',
} as const

export type DocumentKind = (typeof DocumentKind)[keyof typeof DocumentKind]

/**
 * O que o documento tem e o aplicativo não dá conta — em duas categorias que
 * **não** são o mesmo problema (ver docs/02-docx-cirurgico.md).
 *
 * `invisible`: continua no arquivo depois de salvar, mas não aparece na tela.
 * `lost`: some de verdade ao salvar.
 *
 * Misturar os dois produz um aviso genérico que o usuário aprende a ignorar, e
 * aí ele deixa de proteger de qualquer coisa.
 */
export interface LossInventory {
  // Coleções mutáveis, como em `DocumentNode`: este tipo precisa ser atribuível
  // ao que o zod infere no contrato de IPC, e um array `readonly` não é
  // atribuível a um comum. As propriedades continuam `readonly`.
  readonly invisible: string[]
  readonly lost: string[]
}

/**
 * Um arquivo carregado do disco.
 *
 * `content` é sempre texto: `.txt` vem cru e `.sdoc` vem como JSON. Um `.docx`
 * chega aqui **já convertido** para o formato interno pelo processo main, de
 * modo que o renderer segue com um caminho só. Os bytes originais ficam no
 * main, que é quem precisa deles para gravar cirurgicamente.
 */
export interface LoadedFile {
  readonly path: string
  readonly name: string
  readonly kind: DocumentKind
  readonly content: string
  /** Presente só quando o arquivo veio de um formato do Office. */
  readonly inventory?: LossInventory
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
  Open: 'open',
  OpenRecent: 'open-recent',
  ClearRecent: 'clear-recent',
  Save: 'save',
  SaveAs: 'save-as',
  CloseFile: 'close-file',
  FindReplace: 'find-replace',
  ExportPdf: 'export-pdf',
  Print: 'print',
  PrintPreview: 'print-preview',
  PageSetup: 'page-setup',
  InsertPageBreak: 'insert-page-break',
  /** Emitido quando o usuário escolhe "Salvar" no aviso de saída. */
  SaveAndExit: 'save-and-exit',
} as const

export type MenuCommand = (typeof MenuCommand)[keyof typeof MenuCommand]

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
