import { dialog, type BrowserWindow } from 'electron'
import { DiscardChoice, DocumentKind, PlainTextChoice } from '@shared/types.js'
import {
  DOCUMENT_EXTENSION,
  EXCEL_EXTENSION,
  PLAIN_TEXT_EXTENSION,
  SPREADSHEET_EXTENSION,
  OPENABLE_EXTENSIONS,
  PDF_EXTENSION,
  WORD_EXTENSION,
  WORD_MACRO_TEMPLATE_EXTENSION,
  WORD_TEMPLATE_EXTENSION,
} from '@services/file/formats.js'
import { t } from './i18n.js'

const bare = (extension: string) => extension.replace('.', '')

function getFilters() {
  return [
    { name: t('dialog.filter.allSupported'), extensions: OPENABLE_EXTENSIONS.map(bare) },
    { name: t('dialog.filter.wordDocs'), extensions: [bare(WORD_EXTENSION)] },
    { name: t('dialog.filter.wordTemplates'), extensions: templateExtensions() },
    { name: t('dialog.filter.excelSheets'), extensions: [bare(EXCEL_EXTENSION)] },
    { name: t('dialog.filter.documents'), extensions: [bare(DOCUMENT_EXTENSION)] },
    { name: t('dialog.filter.spreadsheets'), extensions: [bare(SPREADSHEET_EXTENSION)] },
    { name: t('dialog.filter.plainText'), extensions: [bare(PLAIN_TEXT_EXTENSION)] },
    { name: t('dialog.filter.pdfDocs'), extensions: [bare(PDF_EXTENSION)] },
    { name: t('dialog.filter.allFiles'), extensions: ['*'] },
  ]
}

/** Only what works for what is open: the open list would offer a spreadsheet to a document. */
function templateExtensions(): string[] {
  return [bare(WORD_TEMPLATE_EXTENSION), bare(WORD_MACRO_TEMPLATE_EXTENSION)]
}

function getSaveFilters(kind: DocumentKind) {
  return kind === DocumentKind.Document
    ? [
        { name: t('dialog.filter.document'), extensions: [bare(DOCUMENT_EXTENSION)] },
        { name: t('dialog.filter.wordDoc'), extensions: [bare(WORD_EXTENSION)] },
        { name: t('dialog.filter.wordTemplate'), extensions: [bare(WORD_TEMPLATE_EXTENSION)] },
        { name: t('dialog.filter.plainText'), extensions: [bare(PLAIN_TEXT_EXTENSION)] },
      ]
    : [
        { name: t('dialog.filter.spreadsheet'), extensions: [bare(SPREADSHEET_EXTENSION)] },
        { name: t('dialog.filter.excelSheet'), extensions: [bare(EXCEL_EXTENSION)] },
      ]
}

function getImageFilters() {
  return [
    { name: t('dialog.filter.images'), extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp'] },
    { name: t('dialog.filter.allFiles'), extensions: ['*'] },
  ]
}

/** `null` when canceled. */
export async function showOpenFileDialog(window: BrowserWindow): Promise<string | null> {
  const result = await dialog.showOpenDialog(window, {
    title: t('dialog.open.title'),
    properties: ['openFile'],
    filters: getFilters(),
  })
  return result.canceled ? null : (result.filePaths[0] ?? null)
}

export async function showTemplatePickerDialog(
  window: BrowserWindow,
  defaultPath: string,
): Promise<string | null> {
  const result = await dialog.showOpenDialog(window, {
    title: t('dialog.template.title'),
    defaultPath,
    properties: ['openFile'],
    filters: [
      { name: t('dialog.filter.wordTemplates'), extensions: templateExtensions() },
      { name: t('dialog.filter.allFiles'), extensions: ['*'] },
    ],
  })
  return result.canceled ? null : (result.filePaths[0] ?? null)
}

export async function showSaveFileDialog(
  window: BrowserWindow,
  suggestedName: string,
  kind: DocumentKind,
): Promise<string | null> {
  const result = await dialog.showSaveDialog(window, {
    title: t('dialog.save.title'),
    defaultPath: suggestedName,
    filters: getSaveFilters(kind),
    // The system dialog already warns about overwriting.
    properties: ['createDirectory', 'showOverwriteConfirmation'],
  })
  return result.canceled ? null : (result.filePath ?? null)
}

/** "Cancel" answers Esc and "Save" answers Enter: a reflex key press loses no work. */
export async function confirmDiscardChanges(window: BrowserWindow, fileName: string): Promise<DiscardChoice> {
  const { response } = await dialog.showMessageBox(window, {
    type: 'warning',
    title: t('dialog.discard.title'),
    message: t('dialog.discard.message', { fileName }),
    detail: t('dialog.discard.detail'),
    buttons: [t('dialog.discard.save'), t('dialog.discard.dontSave'), t('dialog.discard.cancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  })

  if (response === 0) return DiscardChoice.Save
  if (response === 1) return DiscardChoice.Discard
  return DiscardChoice.Cancel
}

/** "Cancel" answers Esc and Enter: a link in a file opens a browser only on a deliberate choice. */
export async function confirmExternalLink(window: BrowserWindow | null, url: string): Promise<boolean> {
  const options = {
    type: 'question' as const,
    title: t('dialog.externalLink.title'),
    message: t('dialog.externalLink.message'),
    detail: url,
    buttons: [t('dialog.externalLink.open'), t('dialog.discard.cancel')],
    defaultId: 1,
    cancelId: 1,
    noLink: true,
  }
  const { response } =
    window === null ? await dialog.showMessageBox(options) : await dialog.showMessageBox(window, options)
  return response === 0
}

/** Asked once per provider, before the first text leaves this computer. */
export async function confirmAiConsent(
  window: BrowserWindow,
  label: string,
  destination: string | undefined,
): Promise<boolean> {
  const { response } = await dialog.showMessageBox(window, {
    type: 'question',
    title: t('ai.consent.title'),
    message: t('ai.consent.message', { provider: label }),
    detail:
      destination === undefined ? t('ai.consent.detail') : `${t('ai.consent.detail')}\n\n${destination}`,
    buttons: [t('ai.consent.allow'), t('dialog.discard.cancel')],
    defaultId: 1,
    cancelId: 1,
    noLink: true,
  })
  return response === 0
}

export async function showPdfSaveDialog(
  window: BrowserWindow,
  suggestedName: string,
): Promise<string | null> {
  const result = await dialog.showSaveDialog(window, {
    title: t('dialog.pdf.title'),
    defaultPath: suggestedName,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
    properties: ['createDirectory', 'showOverwriteConfirmation'],
  })
  return result.canceled ? null : (result.filePath ?? null)
}

const EXPORT_DIALOGS = {
  html: { title: 'dialog.export.htmlTitle', filter: 'dialog.filter.html', extensions: ['html', 'htm'] },
  markdown: {
    title: 'dialog.export.markdownTitle',
    filter: 'dialog.filter.markdown',
    extensions: ['md', 'markdown'],
  },
  odt: { title: 'dialog.export.odtTitle', filter: 'dialog.filter.odt', extensions: ['odt'] },
} as const

export async function showExportSaveDialog(
  window: BrowserWindow,
  suggestedName: string,
  format: keyof typeof EXPORT_DIALOGS,
): Promise<string | null> {
  const chosen = EXPORT_DIALOGS[format]
  const result = await dialog.showSaveDialog(window, {
    title: t(chosen.title),
    defaultPath: suggestedName,
    filters: [{ name: t(chosen.filter), extensions: [...chosen.extensions] }],
    properties: ['createDirectory', 'showOverwriteConfirmation'],
  })
  return result.canceled ? null : (result.filePath ?? null)
}

/** The program the person picks is what runs: a path never arrives from the renderer. */
export async function showAgentExecutableDialog(window: BrowserWindow): Promise<string | null> {
  const result = await dialog.showOpenDialog(window, {
    title: t('dialog.agent.title'),
    properties: ['openFile'],
    ...(process.platform === 'win32'
      ? { filters: [{ name: t('dialog.filter.programs'), extensions: ['exe'] }] }
      : {}),
  })
  return result.canceled ? null : (result.filePaths[0] ?? null)
}

export async function showImagePickerDialog(window: BrowserWindow): Promise<string | null> {
  const result = await dialog.showOpenDialog(window, {
    title: t('dialog.image.title'),
    properties: ['openFile'],
    filters: getImageFilters(),
  })
  return result.canceled ? null : (result.filePaths[0] ?? null)
}

/** The default is saving as a document. */
export async function confirmPlainTextSave(
  window: BrowserWindow,
  fileName: string,
): Promise<PlainTextChoice> {
  const { response } = await dialog.showMessageBox(window, {
    type: 'warning',
    title: t('dialog.plainText.title'),
    message: t('dialog.plainText.message', { fileName }),
    detail: t('dialog.plainText.detail'),
    buttons: [
      t('dialog.plainText.saveAsDocument'),
      t('dialog.plainText.saveAsPlain'),
      t('dialog.discard.cancel'),
    ],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  })

  if (response === 0) return PlainTextChoice.SaveAsDocument
  if (response === 1) return PlainTextChoice.KeepPlain
  return PlainTextChoice.Cancel
}

export function showAboutDialog(window: BrowserWindow, appName: string, version: string): void {
  void dialog.showMessageBox(window, {
    type: 'info',
    title: t('dialog.about.title', { app: appName }),
    message: appName,
    detail: t('dialog.about.detail', { version }),
    buttons: [t('dialog.about.close')],
    noLink: true,
  })
}
