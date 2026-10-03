import { BrowserWindow, type IpcMainInvokeEvent } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { LoadedFile } from '@shared/types.js'
import {
  ensureSupportedExtension,
  fileNameFromPath,
  isExcelPath,
  isWordPackagePath,
  isWordTemplatePath,
  kindFromPath,
  WORD_EXTENSION,
} from '@services/file/formats.js'
import { showOpenFileDialog, showSaveFileDialog } from '../dialogs.js'
import { followDocxOriginal, forgetOpenedDocx, openDocx, saveDocx } from '../docx/index.js'
import { forgetOpenedXlsx, openXlsx, saveXlsx } from '../xlsx/index.js'
import { sidecar } from '../sidecar/index.js'
import { writeFileAtomic } from '../fs/atomic-write.js'
import { assertPathAuthorized, assertReadableFile, authorizePath, normalizePath } from '../fs/paths.js'
import { clearRecentFiles, isRemembered, listRecentFiles, rememberRecentFile } from '../fs/recent.js'
import { readTextFile } from '../fs/read-text.js'
import { refreshMenu } from '../menu.js'
import { t } from '../i18n.js'
import { handle } from './registry.js'
import { forgetExternalFileRequest, isExternalFileRequested } from '../external-files.js'

function windowOf(event: IpcMainInvokeEvent): BrowserWindow {
  const window = BrowserWindow.fromWebContents(event.sender)
  if (window === null) {
    throw new AppError(ErrorCode.Internal, t('errors.ipc.windowNotAvailable'))
  }
  return window
}

/**
 * O modelo do Word abre como documento novo, e o caminho dele **não** é
 * autorizado: nada grava por cima do modelo. `templateName` é o nome do
 * documento novo — o do modelo embutido, traduzido; ausente, o do arquivo.
 */
export async function loadFile(path: string, templateName?: string): Promise<LoadedFile> {
  await assertReadableFile(path)

  // DOCX e XLSX viram formato interno aqui: o renderer nunca vê OOXML.
  const loaded = isWordPackagePath(path)
    ? await openDocx(sidecar(), path)
    : isExcelPath(path)
      ? await openXlsx(sidecar(), path)
      : { content: await readTextFile(path), inventory: undefined }

  // Os bytes originais valem para **um** arquivo: senão a próxima gravação escreveria no pacote errado.
  if (!isWordPackagePath(path)) forgetOpenedDocx()
  if (!isExcelPath(path)) forgetOpenedXlsx()

  if (isWordTemplatePath(path)) {
    const source = normalizePath(path)
    rememberRecentFile(source)
    void refreshMenu()
    const base = templateName ?? fileNameFromPath(source).replace(/\.dot[xm]$/i, '')
    const file: LoadedFile = {
      path: source,
      name: `${base}${WORD_EXTENSION}`,
      kind: kindFromPath(source),
      content: loaded.content,
      template: true,
    }
    return loaded.inventory === undefined ? file : { ...file, inventory: loaded.inventory }
  }

  const authorized = authorizePath(path)
  rememberRecentFile(authorized)
  void refreshMenu()

  const file: LoadedFile = {
    path: authorized,
    name: fileNameFromPath(authorized),
    kind: kindFromPath(authorized),
    content: loaded.content,
  }

  return loaded.inventory === undefined ? file : { ...file, inventory: loaded.inventory }
}

export function registerFileHandlers(): void {
  handle(IpcChannel.FileOpen, async (_payload, event) => {
    const path = await showOpenFileDialog(windowOf(event))
    if (path === null) return { canceled: true as const }
    return { canceled: false as const, file: await loadFile(path) }
  })

  handle(IpcChannel.FileOpenRecent, async (payload) => {
    // O renderer não escolhe caminhos: só recentes ou pedidos do Explorer recebidos pelo main.
    if (!isRemembered(payload.path) && !isExternalFileRequested(payload.path)) {
      throw new AppError(ErrorCode.PathNotAuthorized, t('errors.ipc.notInRecents'))
    }
    try {
      return { file: await loadFile(payload.path) }
    } finally {
      forgetExternalFileRequest(payload.path)
    }
  })

  handle(IpcChannel.FileSave, async (payload) => {
    const path = assertPathAuthorized(payload.path)

    // `.docx` e `.xlsx` vão ao sidecar, que reescreve só o que foi tocado.
    const word = isWordPackagePath(path)
      ? await saveDocx(sidecar(), payload.content, {
          origin: payload.origin,
          destination: path,
          template: isWordTemplatePath(path),
        })
      : null
    const saved = word ?? (isExcelPath(path) ? await saveXlsx(sidecar(), payload.content) : null)
    await writeFileAtomic(path, saved?.bytes ?? payload.content)
    followDocxOriginal(payload.origin, path, word)

    rememberRecentFile(path)
    void refreshMenu()

    const result = { path, name: fileNameFromPath(path) }
    return saved === null ? result : { ...result, inventory: saved.inventory }
  })

  handle(IpcChannel.FileChooseSavePath, async (payload, event) => {
    const chosen = await showSaveFileDialog(windowOf(event), payload.suggestedName, payload.kind)
    if (chosen === null) return { canceled: true as const }

    // Só autoriza: a gravação é outra chamada, para o aviso de perda vir antes de escrever.
    const path = authorizePath(ensureSupportedExtension(chosen, payload.kind))
    return { canceled: false as const, path, name: fileNameFromPath(path) }
  })

  handle(IpcChannel.RecentList, async () => ({ files: [...(await listRecentFiles())] }))

  handle(IpcChannel.RecentClear, async () => {
    clearRecentFiles()
    await refreshMenu()
    return { files: [] }
  })
}
