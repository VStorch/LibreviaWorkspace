import { shell, type Session, type WebContents } from 'electron'
import {
  buildContentSecurityPolicy,
  isAllowedExternalUrl,
  isAllowedNavigation,
  type AppMode,
} from './security-policy.js'

/** Injeta a CSP e nega toda permissão de dispositivo — o app não precisa de nenhuma. */
export function applySessionPolicy(session: Session, mode: AppMode): void {
  const csp = buildContentSecurityPolicy(mode)

  session.webRequest.onHeadersReceived((details, callback) => {
    callback({
      responseHeaders: {
        ...details.responseHeaders,
        'Content-Security-Policy': [csp],
      },
    })
  })

  // Câmera, microfone, geolocalização, notificações: nada disso tem uso aqui.
  session.setPermissionRequestHandler((_wc, _permission, callback) => callback(false))
  session.setPermissionCheckHandler(() => false)
}

/**
 * Trava a navegação da janela e o comportamento de links.
 *
 * `appOrigin` é a URL do servidor de desenvolvimento, ou `null` em produção
 * (onde a interface vem de `file:`).
 */
export function applyNavigationPolicy(contents: WebContents, appOrigin: string | null): void {
  contents.on('will-navigate', (event, url) => {
    if (!isAllowedNavigation(url, appOrigin)) event.preventDefault()
  })

  contents.on('will-attach-webview', (event) => event.preventDefault())

  // Nenhuma janela filha é aberta pelo Chromium. Links legítimos vão para o
  // navegador do sistema, e só depois de passarem pela allowlist de esquema.
  contents.setWindowOpenHandler(({ url }) => {
    if (isAllowedExternalUrl(url)) void shell.openExternal(url)
    return { action: 'deny' }
  })
}
