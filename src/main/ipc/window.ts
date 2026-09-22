import { BrowserWindow, type IpcMainInvokeEvent } from 'electron'
import { AppError, ErrorCode } from '@shared/errors.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { fileNameFromPath } from '@services/file/formats.js'
import { confirmDiscardChanges, confirmPlainTextSave, showImagePickerDialog } from '../dialogs.js'
import { readImageAsDataUrl } from '../fs/read-image.js'
import { listInstalledFontFamilies } from '../system-fonts.js'
import { closeWithoutGuard, updateWindowState } from '../window.js'
import { t } from '../i18n.js'
import { handle } from './registry.js'
import { externalFilesReady } from '../external-files.js'

function windowOf(event: IpcMainInvokeEvent): BrowserWindow {
  const window = BrowserWindow.fromWebContents(event.sender)
  if (window === null) {
    throw new AppError(ErrorCode.Internal, t('errors.ipc.windowNotAvailable'))
  }
  return window
}

export function registerWindowHandlers(): void {
  handle(IpcChannel.WindowReady, (_payload, event) => {
    externalFilesReady(windowOf(event))
    return { applied: true as const }
  })
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

  // A lista de fontes do sistema. Fica entre os handlers de janela porque é da
  // mesma natureza: informação do ambiente que só o main alcança, sem nada a ver
  // com arquivo nem com o documento aberto.
  //
  // A lista sai cortada nos limites do contrato — 4000 famílias, cem caracteres
  // cada — em vez de chegar crua ao schema: o registro valida a resposta, e uma
  // máquina de gráfica com um nome de fonte absurdo derrubaria a lista **inteira**
  // por causa de uma entrada. Perder uma família é melhor que perder a lista.
  handle(IpcChannel.FontsList, async () => ({
    families: (await listInstalledFontFamilies()).filter((family) => family.length <= 100).slice(0, 4000),
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
