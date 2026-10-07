import { IpcChannel } from '@shared/ipc-channels.js'
import { assertViewingAllowed } from '../fs/paths.js'
import { attachPdfView, detachPdfView, layoutPdfView } from '../pdf-viewer.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'

export function registerPdfHandlers(): void {
  handle(IpcChannel.PdfAttach, (payload, event) => {
    attachPdfView(windowOf(event), assertViewingAllowed(payload.path), payload.bounds)
    return { applied: true as const }
  })

  handle(IpcChannel.PdfLayout, (payload, event) => {
    layoutPdfView(windowOf(event), payload.bounds)
    return { applied: true as const }
  })

  handle(IpcChannel.PdfDetach, (_payload, event) => {
    detachPdfView(windowOf(event))
    return { applied: true as const }
  })
}
