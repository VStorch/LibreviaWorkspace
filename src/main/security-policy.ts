import { ALLOWED_EXTERNAL_PROTOCOLS } from '@shared/constants.js'

/**
 * Política de segurança em forma pura — sem importar `electron`, para que seja
 * verificável por teste unitário. `security.ts` é quem a aplica.
 */

/**
 * Preferências obrigatórias de toda janela do aplicativo.
 *
 * O critério de aceite da Fase 0 é que o renderer não alcance o Node.js. São
 * estes quatro valores que garantem isso, e há teste travando cada um deles:
 * uma regressão aqui é silenciosa e catastrófica, então não confiamos na
 * revisão humana.
 */
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
 * CSP da janela.
 *
 * `style-src` precisa de 'unsafe-inline' porque React aplica estilos inline; é
 * um relaxamento conhecido e aceito. Em desenvolvimento, o Vite injeta scripts
 * inline e usa WebSocket para HMR — daí a variante mais permissiva, que nunca
 * chega à build de produção.
 */
export function buildContentSecurityPolicy(mode: AppMode): string {
  const directives: Record<string, string> =
    mode === 'development'
      ? {
          'default-src': "'self'",
          'script-src': "'self' 'unsafe-inline'",
          'style-src': "'self' 'unsafe-inline'",
          'img-src': "'self' data: blob:",
          'font-src': "'self' data:",
          'connect-src': "'self' ws://localhost:* http://localhost:*",
        }
      : {
          'default-src': "'self'",
          'script-src': "'self'",
          'style-src': "'self' 'unsafe-inline'",
          'img-src': "'self' data: blob:",
          'font-src': "'self' data:",
          // O aplicativo é offline: nenhuma requisição de rede é legítima.
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

/**
 * Navegação permitida dentro da janela: apenas a própria origem do aplicativo.
 * Qualquer outra coisa — inclusive um link clicado dentro de um documento —
 * é bloqueada, para que a janela nunca deixe de ser o aplicativo.
 */
export function isAllowedNavigation(targetUrl: string, appOrigin: string | null): boolean {
  let target: URL
  try {
    target = new URL(targetUrl)
  } catch {
    return false
  }

  // Produção: a interface é carregada de disco.
  if (target.protocol === 'file:') return appOrigin === null

  if (appOrigin === null) return false
  try {
    return target.origin === new URL(appOrigin).origin
  } catch {
    return false
  }
}
