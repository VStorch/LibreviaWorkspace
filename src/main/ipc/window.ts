import { BrowserWindow, type IpcMainInvokeEvent } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { fileNameFromPath } from '@services/file/formats.js'
import { confirmDiscardChanges, confirmPlainTextSave, showImagePickerDialog } from '../dialogs.js'
import { readImageAsDataUrl } from '../fs/read-image.js'
import { closeWithoutGuard, updateWindowState } from '../window.js'
import { handle } from './registry.js'

function windowOf(event: IpcMainInvokeEvent): BrowserWindow {
  const window = BrowserWindow.fromWebContents(event.sender)
  if (window === null) {
    throw new AppError(ErrorCode.Internal, 'A janela do aplicativo não está disponível.')
  }
  return window
}

export function registerWindowHandlers(): void {
  // Mesmo aviso nativo usado pelo guarda de fechamento da janela, para que
  // fechar o arquivo, abrir outro ou sair pareçam a mesma coisa ao usuário.
  handle(IpcChannel.DialogConfirmDiscard, async (payload, event) => ({
    choice: await confirmDiscardChanges(windowOf(event), payload.fileName),
  }))

  handle(IpcChannel.DialogConfirmPlainText, async (payload, event) => ({
    choice: await confirmPlainTextSave(windowOf(event), payload.fileName),
  }))

  handle(IpcChannel.ImagePick, async (_payload, event) => {
    const path = await showImagePickerDialog(windowOf(event))
    if (path === null) return { canceled: true as const }

    // A validação por assinatura de bytes acontece aqui, no processo main:
    // o renderer só recebe um data URI de formato já confirmado.
    return {
      canceled: false as const,
      dataUrl: await readImageAsDataUrl(path),
      name: fileNameFromPath(path),
    }
  })

  handle(IpcChannel.WindowSetState, (payload, event) => {
    updateWindowState(windowOf(event), payload.title, payload.isDirty)
    return { applied: true as const }
  })

  // Chamado pelo renderer depois que ele já resolveu o que fazer com as
  // alterações pendentes — daí passar por cima do guarda.
  handle(IpcChannel.WindowClose, (_payload, event) => {
    closeWithoutGuard(windowOf(event))
    return { closing: true as const }
  })
}
