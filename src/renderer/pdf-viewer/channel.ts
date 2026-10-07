import { viewerCommandSchema, type ViewerCommand, type ViewerEvent } from '@shared/pdf-viewer-protocol.js'

/** Only the app window that holds the frame speaks to it; what does not fit the protocol is dropped. */
export function listenToApp(handle: (command: ViewerCommand) => void): void {
  window.addEventListener('message', (event) => {
    if (event.source !== window.parent) return
    const parsed = viewerCommandSchema.safeParse(event.data)
    if (parsed.success) handle(parsed.data)
  })
}

/**
 * `'*'`: the parent is always the app's window, whose origin is `file://` in production, which
 * `postMessage` cannot name. The sandbox and main keep the frame from navigating anywhere else.
 */
export function tellApp(event: ViewerEvent): void {
  window.parent.postMessage(event, '*')
}
