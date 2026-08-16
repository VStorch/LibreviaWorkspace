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
  /**
   * Subconjunto de `invisible`: o que some se o bloco que o ancora for editado.
   *
   * Comentário, revisão, nota e campo calculado entram aqui; posicionamento de
   * imagem e decoração, não. A diferença decide se o arquivo abre em somente
   * leitura — travar a edição por perda de aparência travaria o uso do dia a
   * dia, e o usuário aprenderia a liberar sem ler.
   */
  readonly structural: string[]
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

/**
 * O rascunho de recuperação, sem o conteúdo.
 *
 * O aviso precisa dizer de que arquivo veio e de quando é; carregar junto o
 * conteúdo — que pode ter dezenas de megabytes com imagens embutidas — só para
 * decidir se mostra um aviso seria desperdício.
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
