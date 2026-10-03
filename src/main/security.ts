import { shell, type Session, type WebContents } from 'electron'
import {
  buildContentSecurityPolicy,
  isAllowedExternalUrl,
  isAllowedNavigation,
  type AppMode,
} from './security-policy.js'

/** Nega toda permissão de dispositivo: o app não precisa de nenhuma. */
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

  session.setPermissionRequestHandler((_wc, _permission, callback) => callback(false))
  session.setPermissionCheckHandler(() => false)
}

/** `appOrigin` é o servidor de desenvolvimento, ou `null` em produção (`file:`). */
export function applyNavigationPolicy(contents: WebContents, appOrigin: string | null): void {
  contents.on('will-navigate', (event, url) => {
    if (!isAllowedNavigation(url, appOrigin)) event.preventDefault()
  })

  contents.on('will-attach-webview', (event) => event.preventDefault())

  // Nenhuma janela filha: links vão ao navegador do sistema, depois da lista de esquemas.
  contents.setWindowOpenHandler(({ url }) => {
    if (isAllowedExternalUrl(url)) void shell.openExternal(url)
    return { action: 'deny' }
  })
}
