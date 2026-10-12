import { useCallback, useEffect, useRef, useState, type RefObject } from 'react'
import { clampZoom } from '@services/document/zoom.js'
import { PDF_VIEWER_ORIGIN, PDF_VIEWER_PAGE } from '@shared/constants.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { PAGE_WIDTH, viewerEventSchema, type ViewerEvent } from '@shared/pdf-viewer-protocol.js'
import { isPdfZoomInKey } from '@shared/shortcuts.js'
import { MenuCommand } from '@shared/types.js'
import { EditorCommand, onEditorCommand } from '../document/editor-commands.js'
import { useT } from '../i18n.js'
import { currentPreferences, usePreferences } from '../state/preferences.js'
import { runZoomCommand, setFittedZoom } from '../state/zoom.js'
import { PdfFindBar } from './PdfFindBar.js'
import { PdfPasswordDialog } from './PdfPasswordDialog.js'
import { PdfPropertiesDialog } from './PdfPropertiesDialog.js'
import { connectViewer, sendToViewer, usePdfViewer } from './viewer-state.js'

type Failure = Extract<ViewerEvent, { type: 'error' }>['reason']

interface PasswordRequest {
  readonly retry: boolean
}

const FAILURE_MESSAGES: Readonly<Record<Failure, MessageKey>> = {
  password: 'shell.pdf.password',
  invalid: 'shell.pdf.invalid',
  unknown: 'shell.pdf.failed',
}

/**
 * Its own origin and process: the frame cannot reach `window.api` or the app's page, opens no
 * window, and does not navigate the app away.
 */
const FRAME_SANDBOX = 'allow-scripts allow-same-origin'

/** pdf.js draws the PDF in an isolated frame; the app sends it commands and hears its state. */
export function PdfView({ path }: { readonly path: string }): React.JSX.Element {
  const t = useT()
  const frame = useRef<HTMLIFrameElement>(null)
  const [failure, setFailure] = useState<Failure | null>(null)
  const [passwordRequest, setPasswordRequest] = useState<PasswordRequest | null>(null)
  const ready = useViewerChannel(frame, path, setFailure, setPasswordRequest)
  const [finding, setFinding] = useState(false)
  const [showProperties, setShowProperties] = useState(false)
  const closeProperties = useCallback(() => setShowProperties(false), [])
  useViewerZoom(ready)
  useViewerTheme(ready)
  useZoomInKeys()

  useEffect(
    () =>
      onEditorCommand((command) => {
        if (command === EditorCommand.FindReplace) setFinding(true)
        if (command === EditorCommand.DocumentProperties) setShowProperties(true)
      }),
    [],
  )

  if (failure !== null) {
    return (
      <div className="pdf-view pdf-view--failed" role="alert">
        {t(FAILURE_MESSAGES[failure])}
      </div>
    )
  }

  return (
    <div className="pdf-view">
      {finding && ready && <PdfFindBar onClose={() => setFinding(false)} />}
      {showProperties && <PdfPropertiesDialog path={path} onClose={closeProperties} />}
      {passwordRequest !== null && (
        <PdfPasswordDialog
          retry={passwordRequest.retry}
          onAnswer={(password) => {
            setPasswordRequest(null)
            sendToViewer({ type: 'password', password })
          }}
          onCancel={() => setFailure('password')}
        />
      )}
      <iframe
        ref={frame}
        className="pdf-view__frame"
        src={PDF_VIEWER_PAGE}
        sandbox={FRAME_SANDBOX}
        title={t('shell.pdf.viewer')}
      />
    </div>
  )
}

/** `true` once the frame has the PDF's address: from then on it takes commands. */
function useViewerChannel(
  frame: RefObject<HTMLIFrameElement | null>,
  path: string,
  onFailure: (failure: Failure) => void,
  onPassword: (request: PasswordRequest) => void,
): boolean {
  const [ready, setReady] = useState(false)

  useEffect(() => {
    const target = frame.current?.contentWindow ?? null
    if (target === null) return undefined

    let disposed = false
    const disconnect = connectViewer((command) => target.postMessage(command, PDF_VIEWER_ORIGIN))

    const open = async (): Promise<void> => {
      const shared = await window.api.pdf.share({ path })
      if (disposed) return
      if (!shared.ok) return onFailure('unknown')
      sendToViewer({ type: 'open', url: shared.data.url })
      setReady(true)
    }

    const onMessage = (event: MessageEvent): void => {
      if (event.source !== target || event.origin !== PDF_VIEWER_ORIGIN) return
      const parsed = viewerEventSchema.safeParse(event.data)
      if (!parsed.success) return
      if (parsed.data.type === 'ready') void open()
      else hearViewer(parsed.data, onFailure, onPassword)
    }

    window.addEventListener('message', onMessage)
    return () => {
      disposed = true
      window.removeEventListener('message', onMessage)
      disconnect()
    }
  }, [frame, path, onFailure, onPassword])

  return ready
}

function hearViewer(
  event: Exclude<ViewerEvent, { type: 'ready' }>,
  onFailure: (failure: Failure) => void,
  onPassword: (request: PasswordRequest) => void,
): void {
  switch (event.type) {
    case 'loaded':
      usePdfViewer.setState({ page: 1, pages: event.pages })
      return
    case 'page':
      usePdfViewer.setState({ page: event.page })
      return
    case 'zoom':
      // The viewer measured "fit width"; the status bar shows it as the editor's would.
      if (currentPreferences().zoomFit) setFittedZoom(event.zoom)
      return
    case 'zoom-in':
      void runZoomCommand(MenuCommand.ZoomIn)
      return
    case 'found':
      usePdfViewer.setState({ found: { current: event.current, total: event.total } })
      return
    case 'password':
      onPassword({ retry: event.retry })
      return
    case 'link':
      // Main asks before opening the browser.
      window.open(event.url)
      return
    case 'error':
      onFailure(event.reason)
  }
}

/** The editor's zoom: one preference for every kind of file. */
function useViewerZoom(ready: boolean): void {
  const zoom = usePreferences((state) => state.preferences.zoom)
  const zoomFit = usePreferences((state) => state.preferences.zoomFit)

  useEffect(() => {
    if (ready) sendToViewer({ type: 'zoom', zoom: zoomFit ? PAGE_WIDTH : clampZoom(zoom) })
  }, [ready, zoom, zoomFit])
}

/** With the focus outside the frame, as on the find bar; inside it, the frame asks for the step. */
function useZoomInKeys(): void {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent): void => {
      if (!isPdfZoomInKey(event)) return
      event.preventDefault()
      void runZoomCommand(MenuCommand.ZoomIn)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])
}

/** Follows `data-theme`, which also changes when the system's theme does. */
function useViewerTheme(ready: boolean): void {
  useEffect(() => {
    if (!ready) return undefined
    const root = document.documentElement
    const send = (): void =>
      sendToViewer({ type: 'theme', theme: root.dataset['theme'] === 'dark' ? 'dark' : 'light' })

    send()
    const observer = new MutationObserver(send)
    observer.observe(root, { attributes: true, attributeFilter: ['data-theme'] })
    return () => observer.disconnect()
  }, [ready])
}
