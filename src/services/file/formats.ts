import { DocumentKind, SaveFormat } from '@shared/types.js'
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
/**
 * A destination for any document: the sidecar writes into the `.odt` it came from, or into a
 * minimal package for any other.
 */
export const ODT_EXTENSION = '.odt'
export const ODT_TEMPLATE_EXTENSION = '.ott'
export const SUPPORTED_EXTENSIONS = [
  DOCUMENT_EXTENSION,
  SPREADSHEET_EXTENSION,
  WORD_EXTENSION,
  WORD_TEMPLATE_EXTENSION,
  WORD_MACRO_TEMPLATE_EXTENSION,
  EXCEL_EXTENSION,
  ODT_EXTENSION,
  ODT_TEMPLATE_EXTENSION,
  PLAIN_TEXT_EXTENSION,
] as const
/** Only viewed, never a save destination: that is why it stays out of `SUPPORTED_EXTENSIONS`. */
export const PDF_EXTENSION = '.pdf'
export const OPENABLE_EXTENSIONS = [...SUPPORTED_EXTENSIONS, PDF_EXTENSION] as const

export function isPlainTextPath(path: string): boolean {
  return extensionOf(path) === PLAIN_TEXT_EXTENSION
}

export function isPdfPath(path: string): boolean {
  return extensionOf(path) === PDF_EXTENSION
}

export function isOdtPath(path: string): boolean {
  return extensionOf(path) === ODT_EXTENSION
}

/** Opening creates a new document from the template, as for a Word template. */
export function isOdtTemplatePath(path: string): boolean {
  return extensionOf(path) === ODT_TEMPLATE_EXTENSION
}

/** Any ODF text package the sidecar reads and writes: document or template. */
export function isOdtPackagePath(path: string): boolean {
  return isOdtPath(path) || isOdtTemplatePath(path)
}

/** What the personal data inspector reads: an Office package or an ODF text package. */
export function isInspectablePackagePath(path: string): boolean {
  return isOfficePackagePath(path) || isOdtPackagePath(path)
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

/** A DOCX, Word template or XLSX: an Office package the sidecar opens part by part. */
export function isOfficePackagePath(path: string): boolean {
  return isWordPackagePath(path) || isExcelPath(path)
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

/** What main reads into the editor, and saves to. */
export function isSupportedExtension(path: string): boolean {
  return (SUPPORTED_EXTENSIONS as readonly string[]).includes(extensionOf(path))
}

export function isOpenableExtension(path: string): boolean {
  return (OPENABLE_EXTENSIONS as readonly string[]).includes(extensionOf(path))
}

export function kindFromPath(path: string): DocumentKind {
  const extension = extensionOf(path)
  if (extension === PDF_EXTENSION) return DocumentKind.Pdf
  return extension === SPREADSHEET_EXTENSION || extension === EXCEL_EXTENSION
    ? DocumentKind.Spreadsheet
    : DocumentKind.Document
}

/** Writing text into a file named `.docx` would make Word refuse it. */
export function ensureSupportedExtension(
  path: string,
  kind: DocumentKind = DocumentKind.Document,
  fallback: string = preferredExtension(kind, SaveFormat.Ask),
): string {
  // `.dotm` is not a destination: it would go out without the macros its name promises.
  if (isSupportedExtension(path) && extensionOf(path) !== WORD_MACRO_TEMPLATE_EXTENSION) return path
  return `${path}${fallback}`
}

const DOCUMENT_EXTENSION_FOR: Record<SaveFormat, string> = {
  [SaveFormat.Ask]: DOCUMENT_EXTENSION,
  [SaveFormat.Office]: WORD_EXTENSION,
  [SaveFormat.LibreOffice]: ODT_EXTENSION,
  [SaveFormat.Librevia]: DOCUMENT_EXTENSION,
}

/** `.ods` is not written yet, so LibreOffice spreadsheets go out as `.xlsx`, which Calc opens. */
const SPREADSHEET_EXTENSION_FOR: Record<SaveFormat, string> = {
  [SaveFormat.Ask]: SPREADSHEET_EXTENSION,
  [SaveFormat.Office]: EXCEL_EXTENSION,
  [SaveFormat.LibreOffice]: EXCEL_EXTENSION,
  [SaveFormat.Librevia]: SPREADSHEET_EXTENSION,
}

export function preferredExtension(kind: DocumentKind, format: SaveFormat): string {
  return kind === DocumentKind.Spreadsheet
    ? SPREADSHEET_EXTENSION_FOR[format]
    : DOCUMENT_EXTENSION_FOR[format]
}

/** Filter extensions come without the dot, as Electron takes them. */
export function filtersStartingWith<T extends { readonly extensions: readonly string[] }>(
  filters: readonly T[],
  name: string,
): T[] {
  const extension = extensionOf(name).slice(1)
  const matching = filters.filter((filter) => filter.extensions.includes(extension))
  return [...matching, ...filters.filter((filter) => !matching.includes(filter))]
}

/** `extension` with its dot. */
export function withExtension(name: string, extension: string): string {
  const dot = name.lastIndexOf('.')
  return `${dot > 0 ? name.slice(0, dot) : name}${extension}`
}

export function defaultFileName(
  kind: DocumentKind,
  language: Language = Language.Portuguese,
  format: SaveFormat = SaveFormat.Ask,
): string {
  const name =
    kind === DocumentKind.Spreadsheet
      ? translate(language, 'shell.file.untitledSpreadsheet')
      : translate(language, 'shell.file.untitledDocument')
  return `${name}${preferredExtension(kind, format)}`
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
