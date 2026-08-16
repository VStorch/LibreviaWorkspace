import { DocumentKind } from '@shared/types.js'

/**
 * Conhecimento sobre formatos e caminhos, em forma pura.
 *
 * Esta camada não pode usar `node:path` (o linter impede): ela também roda no
 * renderer. As funções abaixo tratam separadores de Windows e de POSIX, porque
 * o mesmo caminho pode vir de um diálogo nativo ou de uma pasta de rede.
 */

/**
 * Formatos que o aplicativo sabe abrir e salvar.
 *
 * `.sdoc` e `.ssheet` são os formatos internos: guardam o modelo inteiro, sem
 * perda nenhuma. `.docx` e `.xlsx` são os do Office, gravados por cima do
 * arquivo original. `.txt` continua de primeira classe, mas só carrega texto —
 * salvar nele descarta formatação, e por isso o aplicativo avisa antes.
 */
export const DOCUMENT_EXTENSION = '.sdoc'
export const SPREADSHEET_EXTENSION = '.ssheet'
export const PLAIN_TEXT_EXTENSION = '.txt'
export const WORD_EXTENSION = '.docx'
export const EXCEL_EXTENSION = '.xlsx'
export const SUPPORTED_EXTENSIONS = [
  DOCUMENT_EXTENSION,
  SPREADSHEET_EXTENSION,
  WORD_EXTENSION,
  EXCEL_EXTENSION,
  PLAIN_TEXT_EXTENSION,
] as const

export function isPlainTextPath(path: string): boolean {
  return extensionOf(path) === PLAIN_TEXT_EXTENSION
}

export function isWordPath(path: string): boolean {
  return extensionOf(path) === WORD_EXTENSION
}

export function isSpreadsheetPath(path: string): boolean {
  return extensionOf(path) === SPREADSHEET_EXTENSION
}

export function isExcelPath(path: string): boolean {
  return extensionOf(path) === EXCEL_EXTENSION
}

export function extensionOf(path: string): string {
  const name = fileNameFromPath(path)
  const dot = name.lastIndexOf('.')
  if (dot <= 0) return ''
  return name.slice(dot).toLowerCase()
}

export function fileNameFromPath(path: string): string {
  const lastSeparator = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return lastSeparator === -1 ? path : path.slice(lastSeparator + 1)
}

export function isSupportedExtension(path: string): boolean {
  return (SUPPORTED_EXTENSIONS as readonly string[]).includes(extensionOf(path))
}

export function kindFromPath(path: string): DocumentKind {
  const extension = extensionOf(path)
  return extension === SPREADSHEET_EXTENSION || extension === EXCEL_EXTENSION
    ? DocumentKind.Spreadsheet
    : DocumentKind.Document
}

/**
 * Garante que o destino tenha extensão que o aplicativo saiba gravar.
 *
 * No diálogo "salvar como" o usuário pode digitar um nome sem extensão, ou com
 * uma que ainda não suportamos. Gravar conteúdo de texto num arquivo chamado
 * `.docx` produziria um arquivo que o Word recusa a abrir — pior que anexar a
 * extensão certa.
 */
export function ensureSupportedExtension(path: string, kind: DocumentKind = DocumentKind.Document): string {
  if (isSupportedExtension(path)) return path
  // A extensão padrão depende do que está sendo salvo: uma planilha gravada
  // como `.sdoc` abriria como documento vazio na próxima vez.
  const fallback = kind === DocumentKind.Spreadsheet ? SPREADSHEET_EXTENSION : DOCUMENT_EXTENSION
  return `${path}${fallback}`
}

export function defaultFileName(kind: DocumentKind): string {
  return kind === DocumentKind.Spreadsheet
    ? `Planilha sem título${SPREADSHEET_EXTENSION}`
    : `Documento sem título${DOCUMENT_EXTENSION}`
}

/**
 * Título da janela: nome do arquivo, marcador de alteração e nome do app.
 * O `•` é o indicador de não salvo — mesma convenção de editores de código.
 */
export function buildWindowTitle(fileName: string | null, isDirty: boolean, appName: string): string {
  const base = fileName ?? 'Sem título'
  return `${isDirty ? '• ' : ''}${base} — ${appName}`
}
