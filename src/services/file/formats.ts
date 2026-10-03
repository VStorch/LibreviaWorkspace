import { DocumentKind } from '@shared/types.js'
import { translate, Language } from '@shared/i18n/index.js'

/** Sem `node:path`, porque também roda no renderer: separadores de Windows e de POSIX. */

/** `.txt` só carrega texto, e salvar nele descarta a formatação, com aviso antes. */
export const DOCUMENT_EXTENSION = '.sdoc'
export const SPREADSHEET_EXTENSION = '.ssheet'
export const PLAIN_TEXT_EXTENSION = '.txt'
export const WORD_EXTENSION = '.docx'
export const EXCEL_EXTENSION = '.xlsx'
/** O `.dotm` só é aberto: as macros dele não viajam, e por isso não é destino. */
export const WORD_TEMPLATE_EXTENSION = '.dotx'
export const WORD_MACRO_TEMPLATE_EXTENSION = '.dotm'
export const SUPPORTED_EXTENSIONS = [
  DOCUMENT_EXTENSION,
  SPREADSHEET_EXTENSION,
  WORD_EXTENSION,
  WORD_TEMPLATE_EXTENSION,
  WORD_MACRO_TEMPLATE_EXTENSION,
  EXCEL_EXTENSION,
  PLAIN_TEXT_EXTENSION,
] as const

export function isPlainTextPath(path: string): boolean {
  return extensionOf(path) === PLAIN_TEXT_EXTENSION
}

export function isWordPath(path: string): boolean {
  return extensionOf(path) === WORD_EXTENSION
}

/** `.dotx` ou `.dotm`: abrir cria um documento novo a partir do modelo. */
export function isWordTemplatePath(path: string): boolean {
  const extension = extensionOf(path)
  return extension === WORD_TEMPLATE_EXTENSION || extension === WORD_MACRO_TEMPLATE_EXTENSION
}

/** Qualquer pacote do Word que o sidecar lê e grava: documento ou modelo. */
export function isWordPackagePath(path: string): boolean {
  return isWordPath(path) || isWordTemplatePath(path)
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

/** Gravar texto num arquivo chamado `.docx` faria o Word o recusar. */
export function ensureSupportedExtension(path: string, kind: DocumentKind = DocumentKind.Document): string {
  // O `.dotm` não é destino: sairia sem as macros que o nome promete.
  if (isSupportedExtension(path) && extensionOf(path) !== WORD_MACRO_TEMPLATE_EXTENSION) return path
  // A extensão padrão depende do que está sendo salvo: uma planilha gravada
  // como `.sdoc` abriria como documento vazio na próxima vez.
  const fallback = kind === DocumentKind.Spreadsheet ? SPREADSHEET_EXTENSION : DOCUMENT_EXTENSION
  return `${path}${fallback}`
}

export function defaultFileName(kind: DocumentKind, language: Language = Language.Portuguese): string {
  return kind === DocumentKind.Spreadsheet
    ? `${translate(language, 'shell.file.untitledSpreadsheet')}${SPREADSHEET_EXTENSION}`
    : `${translate(language, 'shell.file.untitledDocument')}${DOCUMENT_EXTENSION}`
}

/** O `•` marca o não salvo, como nos editores de código. */
export function buildWindowTitle(
  fileName: string | null,
  isDirty: boolean,
  appName: string,
  untitled: string = translate(Language.Portuguese, 'shell.file.untitled'),
): string {
  const base = fileName ?? untitled
  return `${isDirty ? '• ' : ''}${base} — ${appName}`
}
