export const APP_NAME = 'Librevia'

export const WINDOW_DEFAULTS = {
  width: 1280,
  height: 860,
  minWidth: 940,
  minHeight: 600,
} as const

export const ALLOWED_EXTERNAL_PROTOCOLS = ['http:', 'https:', 'mailto:'] as const

/** The PDF viewer's own origin: a frame isolated from the app, with no access to `window.api`. */
export const PDF_VIEWER_SCHEME = 'librevia-viewer'
export const PDF_VIEWER_ORIGIN = `${PDF_VIEWER_SCHEME}://app`
export const PDF_VIEWER_PAGE = `${PDF_VIEWER_ORIGIN}/pdf-viewer/index.html`
