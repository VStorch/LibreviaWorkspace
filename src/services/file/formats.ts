import { DocumentKind } from '@shared/types.js'
import { translate, Language } from '@shared/i18n/index.js'

/** No `node:path`, because it also runs in the renderer: Windows and POSIX separators. */

/** `.txt` only carries text, and saving to it drops formatting, with a warning first. */
export const DOCUMENT_EXTENSION = '.sdoc'
export const SPREADSHEET_EXTENSION = '.ssheet'
export const PLAIN_TEXT_EXTENSION = '.txt'
export const WORD_EXTENSION = '.docx'
export const EXCEL_EXTENSION = '.xlsx'
/** `.dotm` is only opened: its macros do not travel, so it is never a destination. */
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

/** Opening creates a new document from the template. */
export function isWordTemplatePath(path: string): boolean {
  const extension = extensionOf(path)
  return extension === WORD_TEMPLATE_EXTENSION || extension === WORD_MACRO_TEMPLATE_EXTENSION
}

/** Any Word package the sidecar reads and writes: document or template. */
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

/** Writing text into a file named `.docx` would make Word refuse it. */
export function ensureSupportedExtension(path: string, kind: DocumentKind = DocumentKind.Document): string {
  // `.dotm` is not a destination: it would go out without the macros its name promises.
  if (isSupportedExtension(path) && extensionOf(path) !== WORD_MACRO_TEMPLATE_EXTENSION) return path
  // The default extension depends on what is being saved: a spreadsheet saved as `.sdoc` would open
  // as an empty document next time.
  const fallback = kind === DocumentKind.Spreadsheet ? SPREADSHEET_EXTENSION : DOCUMENT_EXTENSION
  return `${path}${fallback}`
}

export function defaultFileName(kind: DocumentKind, language: Language = Language.Portuguese): string {
  return kind === DocumentKind.Spreadsheet
    ? `${translate(language, 'shell.file.untitledSpreadsheet')}${SPREADSHEET_EXTENSION}`
    : `${translate(language, 'shell.file.untitledDocument')}${DOCUMENT_EXTENSION}`
}

/** The `•` marks unsaved, as in code editors. */
export function buildWindowTitle(
  fileName: string | null,
  isDirty: boolean,
  appName: string,
  untitled: string = translate(Language.Portuguese, 'shell.file.untitled'),
): string {
  const base = fileName ?? untitled
  return `${isDirty ? '• ' : ''}${base} — ${appName}`
}
