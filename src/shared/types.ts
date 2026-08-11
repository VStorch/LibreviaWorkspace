/** O que está sendo editado. Guia ícone, filtros de diálogo e editor. */
export const DocumentKind = {
  Document: 'document',
  Spreadsheet: 'spreadsheet',
} as const

export type DocumentKind = (typeof DocumentKind)[keyof typeof DocumentKind]

/** Um arquivo carregado do disco. Na Fase 1 o conteúdo é texto simples. */
export interface LoadedFile {
  readonly path: string
  readonly name: string
  readonly kind: DocumentKind
  readonly content: string
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
