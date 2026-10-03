import { ALLOWED_EXTERNAL_PROTOCOLS } from '@shared/constants.js'

/** Sem `electron`, para testar; `security.ts` a aplica. */

/** O renderer não alcança o Node.js. Há teste travando cada valor: a regressão aqui é silenciosa. */
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
 * `style-src` aceita 'unsafe-inline' porque o React aplica estilo inline. Em
 * desenvolvimento, o Vite injeta script e usa WebSocket para o HMR.
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
          // O aplicativo é offline.
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

/** Um link de documento só vai para o navegador do sistema se passar aqui. */
export function isAllowedExternalUrl(rawUrl: string): boolean {
  let parsed: URL
  try {
    parsed = new URL(rawUrl)
  } catch {
    return false
  }
  return (ALLOWED_EXTERNAL_PROTOCOLS as readonly string[]).includes(parsed.protocol)
}

/** Só a própria origem: um link dentro de um documento não tira a janela do aplicativo. */
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
