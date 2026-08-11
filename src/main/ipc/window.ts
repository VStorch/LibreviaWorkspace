import { BrowserWindow, type IpcMainInvokeEvent } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { confirmDiscardChanges } from '../dialogs.js'
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
