import { describe, expect, it } from 'vitest'
import {
  SECURE_WEB_PREFERENCES,
  buildContentSecurityPolicy,
  isAllowedExternalUrl,
  isAllowedNavigation,
  isPdfViewerUrl,
  isViewerSchemeUrl,
  VIEWER_CONTENT_SECURITY_POLICY,
} from './security-policy.js'

describe('SECURE_WEB_PREFERENCES', () => {
  // This block makes loosening the isolation impossible by accident: changing any of these values
  // breaks the build.
  it.each([
    ['contextIsolation', true],
    ['nodeIntegration', false],
    ['nodeIntegrationInWorker', false],
    ['nodeIntegrationInSubFrames', false],
    ['sandbox', true],
    ['webSecurity', true],
    ['allowRunningInsecureContent', false],
    ['webviewTag', false],
  ])('keeps %s as %s', (key, expected) => {
    expect(SECURE_WEB_PREFERENCES[key as keyof typeof SECURE_WEB_PREFERENCES]).toBe(expected)
  })
})

describe('buildContentSecurityPolicy', () => {
  it('blocks any network connection in production', () => {
    expect(buildContentSecurityPolicy('production')).toContain("connect-src 'none'")
  })

  it('allows no inline script in production', () => {
    expect(buildContentSecurityPolicy('production')).toContain("script-src 'self'")
    expect(buildContentSecurityPolicy('production')).not.toContain("script-src 'self' 'unsafe-inline'")
  })

  it('allows the Vite WebSocket only in development', () => {
    expect(buildContentSecurityPolicy('development')).toContain('ws://localhost:*')
    expect(buildContentSecurityPolicy('production')).not.toContain('ws://')
  })

  it('only the PDF viewer goes into a frame', () => {
    expect(buildContentSecurityPolicy('production')).toContain('frame-src librevia-viewer:;')
    expect(buildContentSecurityPolicy('development')).toContain('frame-src librevia-viewer:;')
  })

  it.each(['object-src', 'base-uri', 'form-action'])('restricts %s in both modes', (directive) => {
    expect(buildContentSecurityPolicy('production')).toContain(`${directive} 'none'`)
    expect(buildContentSecurityPolicy('development')).toContain(`${directive} 'none'`)
  })
})

describe('PDF viewer policy', () => {
  const directive = (name: string): string =>
    VIEWER_CONTENT_SECURITY_POLICY.split('; ').find((entry) => entry.startsWith(`${name} `)) ?? ''

  it('compiles WebAssembly, but accepts neither eval nor inline script', () => {
    expect(directive('script-src')).toBe("script-src librevia-viewer: 'wasm-unsafe-eval'")
  })

  it('fetches nothing outside its own scheme', () => {
    expect(VIEWER_CONTENT_SECURITY_POLICY).toContain("default-src 'none'")
    expect(directive('connect-src')).toBe('connect-src librevia-viewer:')
    expect(directive('worker-src')).toBe('worker-src librevia-viewer:')
  })

  it('recognizes responses of the scheme, which carry their own policy', () => {
    expect(isViewerSchemeUrl('librevia-viewer://app/pdf-viewer/index.html')).toBe(true)
    expect(isViewerSchemeUrl('file:///app/index.html')).toBe(false)
  })
})

describe('isAllowedExternalUrl', () => {
  it.each(['https://exemplo.com', 'http://intranet.local/doc', 'mailto:alguem@exemplo.com'])(
    'accepts %s',
    (url) => {
      expect(isAllowedExternalUrl(url)).toBe(true)
    },
  )

  it.each([
    'file:///etc/passwd',
    'javascript:alert(1)',
    'data:text/html,<script>alert(1)</script>',
    'vbscript:msgbox',
    'not-a-url',
    '',
  ])('recusa %s', (url) => {
    expect(isAllowedExternalUrl(url)).toBe(false)
  })
})

describe('isAllowedNavigation', () => {
  it('allows the own origin of the development server', () => {
    expect(isAllowedNavigation('http://localhost:5173/index.html', 'http://localhost:5173')).toBe(true)
  })

  it('refuses another origin even in development', () => {
    expect(isAllowedNavigation('https://exemplo.com', 'http://localhost:5173')).toBe(false)
  })

  it('allows file: only in production', () => {
    expect(isAllowedNavigation('file:///app/index.html', null)).toBe(true)
    expect(isAllowedNavigation('file:///etc/passwd', 'http://localhost:5173')).toBe(false)
  })

  it('recusa URL malformada', () => {
    expect(isAllowedNavigation('%%%', null)).toBe(false)
  })
})

describe('isPdfViewerUrl', () => {
  it('recognizes the PDF viewer page of Chromium', () => {
    expect(isPdfViewerUrl('chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/index.html')).toBe(true)
  })

  it('exempts no other origin from the CSP', () => {
    // Only the viewer's own page goes without the app's CSP: anything else here would open a hole.
    expect(isPdfViewerUrl('chrome-extension://outraextensaoqualquer/index.html')).toBe(false)
    expect(isPdfViewerUrl('file:///tmp/relatorio.pdf')).toBe(false)
    expect(isPdfViewerUrl('https://mhjfbmdgcfjbbpaeojofohoefgiehjai/index.html')).toBe(false)
    expect(isPdfViewerUrl('%%%')).toBe(false)
  })
})
