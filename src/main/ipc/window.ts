import { IpcChannel } from '@shared/ipc-channels.js'
import { fileNameFromPath, isOdtTemplatePath, isWordTemplatePath } from '@services/file/formats.js'
import { confirmDiscardChanges, confirmPlainTextSave, showImagePickerDialog } from '../dialogs.js'
import { readImageAsDataUrl } from '../fs/read-image.js'
import { listInstalledFontFamilies } from '../system-fonts.js'
import type { BrowserWindow } from 'electron'
import type { IpcRequest } from '@shared/ipc.js'
import { MenuCommand } from '@shared/types.js'
import { openAppWindow } from '../app-window.js'
import { bringToFront, closeWithoutGuard, updateWindowState, windowHolding } from '../window.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'
import { refreshMenu } from '../menu.js'
import { externalFilesReady } from '../external-files.js'
import { MAX_FONT_FAMILIES, MAX_FONT_FAMILY_LENGTH } from '@shared/limits.js'

export function registerWindowHandlers(): void {
  handle(IpcChannel.WindowReady, (_payload, event) => {
    externalFilesReady(windowOf(event))
    return { applied: true as const }
  })
  // The same prompt as the close guard.
  handle(IpcChannel.DialogConfirmDiscard, async (payload, event) => ({
    choice: await confirmDiscardChanges(windowOf(event), payload.fileName),
  }))

  handle(IpcChannel.DialogConfirmPlainText, async (payload, event) => ({
    choice: await confirmPlainTextSave(windowOf(event), payload.fileName),
  }))

  handle(IpcChannel.ImagePick, async (_payload, event) => {
    const path = await showImagePickerDialog(windowOf(event))
    if (path === null) return { canceled: true as const }

    // Byte signature checked here, in main.
    return {
      canceled: false as const,
      dataUrl: await readImageAsDataUrl(path),
      name: fileNameFromPath(path),
    }
  })

  // Cut to the contract limits: one absurd entry would sink the whole list.
  handle(IpcChannel.FontsList, async () => ({
    families: (await listInstalledFontFamilies())
      .filter((family) => family.length <= MAX_FONT_FAMILY_LENGTH)
      .slice(0, MAX_FONT_FAMILIES),
  }))

  handle(IpcChannel.WindowSetState, (payload, event) => {
    if (updateWindowState(windowOf(event), payload)) void refreshMenu()
    return { applied: true as const }
  })

  handle(IpcChannel.WindowOpen, (payload, event) => ({ placed: openElsewhere(payload, windowOf(event)) }))

  // The renderer has already resolved pending changes.
  handle(IpcChannel.WindowClose, (_payload, event) => {
    closeWithoutGuard(windowOf(event))
    return { closing: true as const }
  })
}

type WindowOpening = IpcRequest<typeof IpcChannel.WindowOpen>

const isTemplatePath = (path: string): boolean => isWordTemplatePath(path) || isOdtTemplatePath(path)

/** A template always makes a new document, so the window editing the template itself is no answer. */
function openElsewhere(request: WindowOpening, asking: BrowserWindow): boolean {
  if (request.kind === 'new') return openAppWindow({ command: request.command }) !== null
  if (request.kind === 'template') {
    return openAppWindow({ command: MenuCommand.NewFromTemplate, template: request.template }) !== null
  }
  const holder = isTemplatePath(request.path) ? null : windowHolding(request.path, asking)
  if (holder !== null) {
    bringToFront(holder)
    return true
  }
  if (request.hasRoom) return false
  return openAppWindow({ command: MenuCommand.OpenRecent, path: request.path }) !== null
}
