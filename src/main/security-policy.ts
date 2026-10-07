import { ALLOWED_EXTERNAL_PROTOCOLS, PDF_VIEWER_SCHEME } from '@shared/constants.js'

/** Without `electron`, to be testable; `security.ts` applies it. */

/** The renderer cannot reach Node.js. A test pins each value: a regression here is silent. */
export const SECURE_WEB_PREFERENCES = {
  contextIsolation: true,
  nodeIntegration: false,
  nodeIntegrationInWorker: false,
  nodeIntegrationInSubFrames: false,
  sandbox: true,
  webSecurity: true,
  allowRunningInsecureContent: false,
  experimentalFeatures: false,
  webviewTag: false,
} as const

export type AppMode = 'development' | 'production'

/**
 * `style-src` allows 'unsafe-inline' because React applies inline styles. In development Vite
 * injects scripts and uses a WebSocket for HMR.
 */
export function buildContentSecurityPolicy(mode: AppMode): string {
  const directives: Record<string, string> =
    mode === 'development'
      ? {
          'default-src': "'self'",
          'script-src': "'self' 'unsafe-inline'",
          'style-src': "'self' 'unsafe-inline'",
          'img-src': "'self' data: blob:",
          'font-src': "'self' data: librevia-font:",
          'connect-src': "'self' ws://localhost:* http://localhost:*",
        }
      : {
          'default-src': "'self'",
          'script-src': "'self'",
          'style-src': "'self' 'unsafe-inline'",
          'img-src': "'self' data: blob:",
          'font-src': "'self' data: librevia-font:",
          // The app is offline.
          'connect-src': "'none'",
        }

  const common: Record<string, string> = {
    'object-src': "'none'",
    // Only the PDF viewer, which runs in its own origin and process.
    'frame-src': `${PDF_VIEWER_SCHEME}:`,
    'media-src': "'none'",
    'worker-src': "'self'",
    'base-uri': "'none'",
    'form-action': "'none'",
  }

  return Object.entries({ ...directives, ...common })
    .map(([key, value]) => `${key} ${value}`)
    .join('; ')
}

/**
 * The PDF viewer frame's own policy. `'wasm-unsafe-eval'` compiles the JPEG 2000 and JBIG2 image
 * decoders and nothing else: `eval` and `new Function` stay blocked, which is what stops a font
 * crafted to run code (CVE-2024-4367). pdf.js positions the text layer with inline styles.
 */
export const VIEWER_CONTENT_SECURITY_POLICY = [
  "default-src 'none'",
  `script-src ${PDF_VIEWER_SCHEME}: 'wasm-unsafe-eval'`,
  `worker-src ${PDF_VIEWER_SCHEME}:`,
  `connect-src ${PDF_VIEWER_SCHEME}:`,
  `img-src ${PDF_VIEWER_SCHEME}: blob: data:`,
  `style-src ${PDF_VIEWER_SCHEME}: 'unsafe-inline'`,
  `font-src ${PDF_VIEWER_SCHEME}: data: blob:`,
  "base-uri 'none'",
  "form-action 'none'",
].join('; ')

/** Its responses carry their own policy; the app's would block the viewer. */
export function isViewerSchemeUrl(rawUrl: string): boolean {
  return rawUrl.startsWith(`${PDF_VIEWER_SCHEME}:`)
}

/** A document link only goes to the system browser if it passes here. */
export function isAllowedExternalUrl(rawUrl: string): boolean {
  let parsed: URL
  try {
    parsed = new URL(rawUrl)
  } catch {
    return false
  }
  return (ALLOWED_EXTERNAL_PROTOCOLS as readonly string[]).includes(parsed.protocol)
}

/**
 * Only the app's own origin: a link inside a document does not take the window away from the app.
 */
export function isAllowedNavigation(targetUrl: string, appOrigin: string | null): boolean {
  let target: URL
  try {
    target = new URL(targetUrl)
  } catch {
    return false
  }

  if (target.protocol === 'file:') return appOrigin === null

  if (appOrigin === null) return false
  try {
    return target.origin === new URL(appOrigin).origin
  } catch {
    return false
  }
}

/** Chromium's built-in PDF viewer: an extension compiled into Electron, not something we load. */
const PDF_VIEWER_EXTENSION_ID = 'mhjfbmdgcfjbbpaeojofohoefgiehjai'

/**
 * The app's CSP is made for the interface: on the viewer's page it would block the
 * `chrome://resources` scripts the viewer is built from, and the PDF would open blank. Compares
 * scheme and host because `origin` is `"null"` for a `chrome-extension:` URL.
 */
export function isPdfViewerUrl(rawUrl: string): boolean {
  try {
    const url = new URL(rawUrl)
    return url.protocol === 'chrome-extension:' && url.host === PDF_VIEWER_EXTENSION_ID
  } catch {
    return false
  }
}
