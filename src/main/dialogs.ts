import { dialog, type BrowserWindow } from 'electron'
import { DiscardChoice, DocumentKind, PlainTextChoice } from '@shared/types.js'
import {
  DOCUMENT_EXTENSION,
  EXCEL_EXTENSION,
  PLAIN_TEXT_EXTENSION,
  SPREADSHEET_EXTENSION,
  SUPPORTED_EXTENSIONS,
  WORD_EXTENSION,
} from '@services/file/formats.js'
import { t } from './i18n.js'

const bare = (extension: string) => extension.replace('.', '')

function getFilters() {
  return [
    { name: t('dialog.filter.allSupported'), extensions: SUPPORTED_EXTENSIONS.map(bare) },
    { name: t('dialog.filter.wordDocs'), extensions: [bare(WORD_EXTENSION)] },
    { name: t('dialog.filter.excelSheets'), extensions: [bare(EXCEL_EXTENSION)] },
    { name: t('dialog.filter.documents'), extensions: [bare(DOCUMENT_EXTENSION)] },
    { name: t('dialog.filter.spreadsheets'), extensions: [bare(SPREADSHEET_EXTENSION)] },
    { name: t('dialog.filter.plainText'), extensions: [bare(PLAIN_TEXT_EXTENSION)] },
    { name: t('dialog.filter.allFiles'), extensions: ['*'] },
  ]
}

/**
 * Os formatos em que cada tipo pode ser salvo, o nativo primeiro.
 *
 * Com a lista de abrir, o diálogo de salvar oferecia planilha do Excel a um
 * documento — e o `.docx` de um documento novo era recusado depois da escolha.
 * Agora o documento novo também grava em `.docx`, e a lista oferece só o que
 * de fato funciona para o que está aberto. O `.txt` continua na lista, com o
 * aviso de formatação perdida que vem antes da gravação.
 */
function getSaveFilters(kind: DocumentKind) {
  return kind === DocumentKind.Document
    ? [
        { name: t('dialog.filter.document'), extensions: [bare(DOCUMENT_EXTENSION)] },
        { name: t('dialog.filter.wordDoc'), extensions: [bare(WORD_EXTENSION)] },
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

/** Devolve o caminho escolhido, ou `null` se o usuário cancelou. */
export async function showOpenFileDialog(window: BrowserWindow): Promise<string | null> {
  const result = await dialog.showOpenDialog(window, {
    title: t('dialog.open.title'),
    properties: ['openFile'],
    filters: getFilters(),
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
    // O diálogo do sistema já avisa sobre sobrescrever; não duplicamos o aviso.
    properties: ['createDirectory', 'showOverwriteConfirmation'],
  })
  return result.canceled ? null : (result.filePath ?? null)
}

/**
 * Aviso de alterações não salvas.
 *
 * "Cancelar" é o botão de fuga (Esc) e "Salvar" é o padrão (Enter): quem
 * apertar uma tecla por reflexo não perde trabalho.
 */
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

export async function showImagePickerDialog(window: BrowserWindow): Promise<string | null> {
  const result = await dialog.showOpenDialog(window, {
    title: t('dialog.image.title'),
    properties: ['openFile'],
    filters: getImageFilters(),
  })
  return result.canceled ? null : (result.filePaths[0] ?? null)
}

/**
 * Aviso de que `.txt` não guarda formatação.
 *
 * A alternativa — salvar em silêncio e descartar negrito, títulos, tabelas e
 * imagens — é exatamente o tipo de perda que o plano se compromete a evitar
 * (§6.1). Por isso o padrão oferecido é salvar como documento.
 */
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
