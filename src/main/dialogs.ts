import { dialog, type BrowserWindow } from 'electron'
import { DiscardChoice, PlainTextChoice } from '@shared/types.js'
import { DOCUMENT_EXTENSION, PLAIN_TEXT_EXTENSION } from '@services/file/formats.js'

const FILTERS = [
  { name: 'Documentos', extensions: [DOCUMENT_EXTENSION.replace('.', '')] },
  { name: 'Texto simples', extensions: [PLAIN_TEXT_EXTENSION.replace('.', '')] },
  { name: 'Todos os arquivos', extensions: ['*'] },
]

const IMAGE_FILTERS = [
  { name: 'Imagens', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp'] },
  { name: 'Todos os arquivos', extensions: ['*'] },
]

/** Devolve o caminho escolhido, ou `null` se o usuário cancelou. */
export async function showOpenFileDialog(window: BrowserWindow): Promise<string | null> {
  const result = await dialog.showOpenDialog(window, {
    title: 'Abrir arquivo',
    properties: ['openFile'],
    filters: FILTERS,
  })
  return result.canceled ? null : (result.filePaths[0] ?? null)
}

export async function showSaveFileDialog(
  window: BrowserWindow,
  suggestedName: string,
): Promise<string | null> {
  const result = await dialog.showSaveDialog(window, {
    title: 'Salvar como',
    defaultPath: suggestedName,
    filters: FILTERS,
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
    title: 'Alterações não salvas',
    message: `Salvar as alterações em “${fileName}”?`,
    detail: 'Se não salvar, as alterações feitas desde a última gravação serão perdidas.',
    buttons: ['Salvar', 'Não salvar', 'Cancelar'],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  })

  if (response === 0) return DiscardChoice.Save
  if (response === 1) return DiscardChoice.Discard
  return DiscardChoice.Cancel
}

export async function showImagePickerDialog(window: BrowserWindow): Promise<string | null> {
  const result = await dialog.showOpenDialog(window, {
    title: 'Inserir imagem',
    properties: ['openFile'],
    filters: IMAGE_FILTERS,
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
    title: 'Formatação será perdida',
    message: `“${fileName}” é um arquivo de texto simples.`,
    detail:
      'Texto simples não guarda negrito, títulos, listas, tabelas nem imagens. ' +
      'Salvar como documento preserva tudo.',
    buttons: ['Salvar como documento', 'Salvar como texto simples', 'Cancelar'],
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
    title: `Sobre o ${appName}`,
    message: appName,
    detail: `Versão ${version}\n\nSuíte de documentos e planilhas, offline.\nEm desenvolvimento — Fase 1 de 8.`,
    buttons: ['Fechar'],
    noLink: true,
  })
}
