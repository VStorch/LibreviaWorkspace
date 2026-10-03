import { dialog, type BrowserWindow } from 'electron'
import { DiscardChoice, DocumentKind, PlainTextChoice } from '@shared/types.js'
import {
  DOCUMENT_EXTENSION,
  EXCEL_EXTENSION,
  PLAIN_TEXT_EXTENSION,
  SPREADSHEET_EXTENSION,
  SUPPORTED_EXTENSIONS,
  WORD_EXTENSION,
  WORD_MACRO_TEMPLATE_EXTENSION,
  WORD_TEMPLATE_EXTENSION,
} from '@services/file/formats.js'
import { t } from './i18n.js'

const bare = (extension: string) => extension.replace('.', '')

function getFilters() {
  return [
    { name: t('dialog.filter.allSupported'), extensions: SUPPORTED_EXTENSIONS.map(bare) },
    { name: t('dialog.filter.wordDocs'), extensions: [bare(WORD_EXTENSION)] },
    { name: t('dialog.filter.wordTemplates'), extensions: templateExtensions() },
    { name: t('dialog.filter.excelSheets'), extensions: [bare(EXCEL_EXTENSION)] },
    { name: t('dialog.filter.documents'), extensions: [bare(DOCUMENT_EXTENSION)] },
    { name: t('dialog.filter.spreadsheets'), extensions: [bare(SPREADSHEET_EXTENSION)] },
    { name: t('dialog.filter.plainText'), extensions: [bare(PLAIN_TEXT_EXTENSION)] },
    { name: t('dialog.filter.allFiles'), extensions: ['*'] },
  ]
}

/** Só o que funciona para o que está aberto: a lista de abrir ofereceria planilha a um documento. */
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

/** `null` se a pessoa cancelou. */
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
    // O diálogo do sistema já avisa sobre sobrescrever.
    properties: ['createDirectory', 'showOverwriteConfirmation'],
  })
  return result.canceled ? null : (result.filePath ?? null)
}

/** "Cancelar" responde ao Esc e "Salvar" ao Enter: a tecla apertada por reflexo não perde trabalho. */
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

export async function showImagePickerDialog(window: BrowserWindow): Promise<string | null> {
  const result = await dialog.showOpenDialog(window, {
    title: t('dialog.image.title'),
    properties: ['openFile'],
    filters: getImageFilters(),
  })
  return result.canceled ? null : (result.filePaths[0] ?? null)
}

/** O padrão oferecido é salvar como documento. */
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
