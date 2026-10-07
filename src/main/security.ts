import { shell, type Session, type WebContents } from 'electron'
import {
  buildContentSecurityPolicy,
  isAllowedExternalUrl,
  isAllowedNavigation,
  isPdfViewerUrl,
  type AppMode,
} from './security-policy.js'

/** The app needs no device permissions. */
export function applySessionPolicy(session: Session, mode: AppMode): void {
  const csp = buildContentSecurityPolicy(mode)

  session.webRequest.onHeadersReceived((details, callback) => {
    if (isPdfViewerUrl(details.url)) {
      callback({})
      return
    }
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

/** `appOrigin` is the development server, or `null` in production (`file:`). */
export function applyNavigationPolicy(contents: WebContents, appOrigin: string | null): void {
  contents.on('will-navigate', (event, url) => {
    if (!isAllowedNavigation(url, appOrigin)) event.preventDefault()
  })

  contents.on('will-attach-webview', (event) => event.preventDefault())

  // No child windows: links go to the system browser, after the scheme allowlist.
  contents.setWindowOpenHandler(({ url }) => {
    if (isAllowedExternalUrl(url)) void shell.openExternal(url)
    return { action: 'deny' }
  })
}

/** A PDF viewer never leaves its file: a link goes to the system browser, or nowhere. */
export function applyViewerNavigationPolicy(contents: WebContents): void {
  const openOutside = (url: string): void => {
    if (isAllowedExternalUrl(url)) void shell.openExternal(url)
  }

  contents.on('will-navigate', (event, url) => {
    event.preventDefault()
    openOutside(url)
  })

  contents.on('will-attach-webview', (event) => event.preventDefault())

  contents.setWindowOpenHandler(({ url }) => {
    openOutside(url)
    return { action: 'deny' }
  })
}
