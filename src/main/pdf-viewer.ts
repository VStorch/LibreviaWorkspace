import { BrowserWindow, WebContentsView, type Rectangle, type WebContents } from 'electron'
import { APP_NAME } from '@shared/constants.js'
import { SECURE_WEB_PREFERENCES } from './security-policy.js'
import { applyViewerNavigationPolicy } from './security.js'

/**
 * Chromium's own viewer, the same as the browser's, with search, zoom, thumbnails and printing. It
 * is a web page, so JavaScript stays on; like the editor's, it has no preload and no Node, and the
 * PDF's links go out through the same allowlist.
 */
const PDF_WEB_PREFERENCES = { ...SECURE_WEB_PREFERENCES, plugins: true } as const

const PREVIEW_SIZE = { width: 900, height: 1000 } as const

interface PdfWindowOptions {
  readonly path: string
  readonly title: string
  readonly parent?: BrowserWindow
}

interface AttachedPdf {
  readonly view: WebContentsView
  readonly onClosed: () => void
  path: string
}

const attachedViews = new WeakMap<BrowserWindow, AttachedPdf>()

/** Print preview: a window apart, because it shows paper the document has not become yet. */
export function openPdfWindow({ path, title, parent }: PdfWindowOptions): BrowserWindow {
  const window = new BrowserWindow({
    ...(parent === undefined ? {} : { parent }),
    ...PREVIEW_SIZE,
    title: `${title} — ${APP_NAME}`,
    autoHideMenuBar: true,
    webPreferences: PDF_WEB_PREFERENCES,
  })

  // Chromium would replace the title with the PDF's own, which is often empty or a generator's.
  window.on('page-title-updated', (event) => event.preventDefault())
  loadPdf(window.webContents, path)
  return window
}

/** The viewer goes inside the window, over the area the renderer reserved; one per window. */
export function attachPdfView(window: BrowserWindow, path: string, bounds: Rectangle): void {
  const attached = attachedViews.get(window) ?? createAttachedView(window)
  if (attached.path !== path) {
    attached.path = path
    loadPdf(attached.view.webContents, path)
  }
  layoutPdfView(window, bounds)
}

/** `null` hides it: the view is drawn above the page, and would cover the app's dialogs. */
export function layoutPdfView(window: BrowserWindow, bounds: Rectangle | null): void {
  const attached = attachedViews.get(window)
  if (attached === undefined) return

  if (bounds !== null) attached.view.setBounds(toWindowPixels(window, bounds))
  attached.view.setVisible(bounds !== null)
}

export function detachPdfView(window: BrowserWindow): void {
  const attached = attachedViews.get(window)
  if (attached === undefined) return

  attachedViews.delete(window)
  window.removeListener('closed', attached.onClosed)
  if (!window.isDestroyed()) window.contentView.removeChildView(attached.view)
  // A view's contents outlive the window unless closed.
  attached.view.webContents.close()
}

function createAttachedView(window: BrowserWindow): AttachedPdf {
  const view = new WebContentsView({ webPreferences: PDF_WEB_PREFERENCES })
  applyViewerNavigationPolicy(view.webContents)
  window.contentView.addChildView(view)

  const attached: AttachedPdf = { view, onClosed: () => detachPdfView(window), path: '' }
  window.once('closed', attached.onClosed)
  attachedViews.set(window, attached)
  return attached
}

/** The renderer measures in CSS pixels; the view is placed in the window's, which zoom scales. */
function toWindowPixels(window: BrowserWindow, bounds: Rectangle): Rectangle {
  const zoom = window.webContents.getZoomFactor()
  return {
    x: Math.round(bounds.x * zoom),
    y: Math.round(bounds.y * zoom),
    width: Math.round(bounds.width * zoom),
    height: Math.round(bounds.height * zoom),
  }
}

function loadPdf(contents: WebContents, path: string): void {
  // Closing before the PDF loads rejects the load; any other failure shows in the viewer.
  contents.loadFile(path).catch(() => undefined)
}
