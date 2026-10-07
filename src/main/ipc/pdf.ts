import { fileNameFromPath } from '@services/file/formats.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { assertViewingAllowed } from '../fs/paths.js'
import { t } from '../i18n.js'
import { shareDocument } from '../pdf-scheme.js'
import { openPdfWindow } from '../pdf-viewer.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'

export function registerPdfHandlers(): void {
  handle(IpcChannel.PdfShare, (payload) => ({ url: shareDocument(assertViewingAllowed(payload.path)) }))

  handle(IpcChannel.PdfPrint, (payload, event) => {
    const path = assertViewingAllowed(payload.path)
    openPdfWindow({
      path,
      title: t('errors.print.previewTitle', { title: fileNameFromPath(path) }),
      parent: windowOf(event),
    })
    return { opened: true as const }
  })
}
