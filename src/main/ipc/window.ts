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
import { setRevisionViewChecked, setTrackChangesChecked } from '../menu.js'
import { externalFilesReady } from '../external-files.js'
import { MAX_FONT_FAMILIES, MAX_FONT_FAMILY_LENGTH } from '@shared/limits.js'

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
  // O mesmo aviso do guarda de fechamento.
  handle(IpcChannel.DialogConfirmDiscard, async (payload, event) => ({
    choice: await confirmDiscardChanges(windowOf(event), payload.fileName),
  }))

  handle(IpcChannel.DialogConfirmPlainText, async (payload, event) => ({
    choice: await confirmPlainTextSave(windowOf(event), payload.fileName),
  }))

  handle(IpcChannel.ImagePick, async (_payload, event) => {
    const path = await showImagePickerDialog(windowOf(event))
    if (path === null) return { canceled: true as const }

    // Assinatura de bytes conferida aqui, no main.
    return {
      canceled: false as const,
      dataUrl: await readImageAsDataUrl(path),
      name: fileNameFromPath(path),
    }
  })

  // Cortada nos limites do contrato: uma entrada absurda derrubaria a lista inteira.
  handle(IpcChannel.FontsList, async () => ({
    families: (await listInstalledFontFamilies())
      .filter((family) => family.length <= MAX_FONT_FAMILY_LENGTH)
      .slice(0, MAX_FONT_FAMILIES),
  }))

  handle(IpcChannel.WindowSetState, (payload, event) => {
    updateWindowState(windowOf(event), payload.title, payload.isDirty)
    setTrackChangesChecked(payload.trackChanges)
    setRevisionViewChecked(payload.revisionView)
    return { applied: true as const }
  })

  // O renderer já resolveu as alterações pendentes.
  handle(IpcChannel.WindowClose, (_payload, event) => {
    closeWithoutGuard(windowOf(event))
    return { closing: true as const }
  })
}
