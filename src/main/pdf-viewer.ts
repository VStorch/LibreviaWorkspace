import { BrowserWindow } from 'electron'
import { APP_NAME } from '@shared/constants.js'
import { SECURE_WEB_PREFERENCES } from './security-policy.js'
import { applyViewerNavigationPolicy } from './security.js'

/**
 * Chromium's own viewer, which prints a PDF as it is, without turning pages into images. It is a
 * web page, so JavaScript stays on; like the editor's, it has no preload and no Node, and the PDF's
 * links go out through the same confirmation.
 */
const PDF_WEB_PREFERENCES = { ...SECURE_WEB_PREFERENCES, plugins: true } as const

const PREVIEW_SIZE = { width: 900, height: 1000 } as const

interface PdfWindowOptions {
  readonly path: string
  readonly title: string
  readonly parent?: BrowserWindow
}

/** Print preview: a window apart, with the printing tools the app's viewer leaves out. */
export function openPdfWindow({ path, title, parent }: PdfWindowOptions): BrowserWindow {
  const window = new BrowserWindow({
    ...(parent === undefined ? {} : { parent }),
    ...PREVIEW_SIZE,
    title: `${title} — ${APP_NAME}`,
    autoHideMenuBar: true,
    webPreferences: PDF_WEB_PREFERENCES,
  })
  applyViewerNavigationPolicy(window.webContents)

  // Chromium would replace the title with the PDF's own, which is often empty or a generator's.
  window.on('page-title-updated', (event) => event.preventDefault())
  // Closing before the PDF loads rejects the load; any other failure shows in the viewer.
  window.loadFile(path).catch(() => undefined)
  return window
}
