import { ALLOWED_EXTERNAL_PROTOCOLS } from '@shared/constants.js'

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
    'frame-src': "'none'",
    'media-src': "'none'",
    'worker-src': "'self'",
    'base-uri': "'none'",
    'form-action': "'none'",
  }

  return Object.entries({ ...directives, ...common })
    .map(([key, value]) => `${key} ${value}`)
    .join('; ')
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
