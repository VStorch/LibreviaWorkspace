import { BrowserWindow, type IpcMainInvokeEvent } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import type { LoadedFile } from '@shared/types.js'
import {
  ensureSupportedExtension,
  fileNameFromPath,
  isWordPath,
  kindFromPath,
} from '@services/file/formats.js'
import { showOpenFileDialog, showSaveFileDialog } from '../dialogs.js'
import { forgetOpenedDocx, openDocx, saveDocx } from '../docx/index.js'
import { sidecar } from '../sidecar/index.js'
import { writeFileAtomic } from '../fs/atomic-write.js'
import { assertPathAuthorized, assertReadableFile, authorizePath } from '../fs/paths.js'
import { clearRecentFiles, isRemembered, listRecentFiles, rememberRecentFile } from '../fs/recent.js'
import { readTextFile } from '../fs/read-text.js'
import { refreshMenu } from '../menu.js'
import { handle } from './registry.js'

function windowOf(event: IpcMainInvokeEvent): BrowserWindow {
  const window = BrowserWindow.fromWebContents(event.sender)
  if (window === null) {
    throw new AppError(ErrorCode.Internal, 'A janela do aplicativo não está disponível.')
  }
  return window
}

/** Valida, lê e passa a considerar o caminho autorizado para gravação. */
async function loadFile(path: string): Promise<LoadedFile> {
  await assertReadableFile(path)

  // O DOCX vira formato interno aqui, e não no renderer: assim o renderer
  // segue com um caminho só e nunca vê OOXML.
  const loaded = isWordPath(path)
    ? await openDocx(sidecar(), path)
    : { content: await readTextFile(path), inventory: undefined }

  if (!isWordPath(path)) forgetOpenedDocx()

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
    // O renderer não escolhe caminhos: só pode reabrir o que já está na lista
    // de recentes, que por sua vez só é alimentada por escolha do usuário.
    if (!isRemembered(payload.path)) {
      throw new AppError(
        ErrorCode.PathNotAuthorized,
        'Este arquivo não está mais na lista de recentes. Abra-o novamente pelo menu Arquivo.',
      )
    }
    return { file: await loadFile(payload.path) }
  })

  handle(IpcChannel.FileSave, async (payload) => {
    const path = assertPathAuthorized(payload.path)

    // Gravar `.docx` não escreve o que o renderer mandou: manda o modelo ao
    // sidecar, que reescreve só os blocos tocados sobre o pacote original.
    const saved = isWordPath(path) ? await saveDocx(sidecar(), payload.content) : null
    await writeFileAtomic(path, saved?.bytes ?? payload.content)

    rememberRecentFile(path)
    void refreshMenu()

    const result = { path, name: fileNameFromPath(path) }
    return saved === null ? result : { ...result, inventory: saved.inventory }
  })

  handle(IpcChannel.FileChooseSavePath, async (payload, event) => {
    const chosen = await showSaveFileDialog(windowOf(event), payload.suggestedName)
    if (chosen === null) return { canceled: true as const }

    // Só autoriza o destino. A gravação é uma chamada separada, para que o
    // renderer possa avisar sobre perda de formatação antes de escrever —
    // e para que cancelar esse aviso não deixe um arquivo pela metade.
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
