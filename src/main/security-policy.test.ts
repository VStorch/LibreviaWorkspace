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
  ])('mantém %s como %s', (key, expected) => {
    expect(SECURE_WEB_PREFERENCES[key as keyof typeof SECURE_WEB_PREFERENCES]).toBe(expected)
  })
})

describe('buildContentSecurityPolicy', () => {
  it('bloqueia qualquer conexão de rede em produção', () => {
    expect(buildContentSecurityPolicy('production')).toContain("connect-src 'none'")
  })

  it('não permite script inline em produção', () => {
    expect(buildContentSecurityPolicy('production')).toContain("script-src 'self'")
    expect(buildContentSecurityPolicy('production')).not.toContain("script-src 'self' 'unsafe-inline'")
  })

  it('libera o WebSocket do Vite apenas em desenvolvimento', () => {
    expect(buildContentSecurityPolicy('development')).toContain('ws://localhost:*')
    expect(buildContentSecurityPolicy('production')).not.toContain('ws://')
  })

  it('só o visualizador de PDF entra num frame', () => {
    expect(buildContentSecurityPolicy('production')).toContain('frame-src librevia-viewer:;')
    expect(buildContentSecurityPolicy('development')).toContain('frame-src librevia-viewer:;')
  })

  it.each(['object-src', 'base-uri', 'form-action'])('restringe %s nos dois modos', (directive) => {
    expect(buildContentSecurityPolicy('production')).toContain(`${directive} 'none'`)
    expect(buildContentSecurityPolicy('development')).toContain(`${directive} 'none'`)
  })
})

describe('política do visualizador de PDF', () => {
  const directive = (name: string): string =>
    VIEWER_CONTENT_SECURITY_POLICY.split('; ').find((entry) => entry.startsWith(`${name} `)) ?? ''

  it('compila WebAssembly, mas não aceita eval nem script inline', () => {
    expect(directive('script-src')).toBe("script-src librevia-viewer: 'wasm-unsafe-eval'")
  })

  it('não busca nada fora do próprio esquema', () => {
    expect(VIEWER_CONTENT_SECURITY_POLICY).toContain("default-src 'none'")
    expect(directive('connect-src')).toBe('connect-src librevia-viewer:')
    expect(directive('worker-src')).toBe('worker-src librevia-viewer:')
  })

  it('reconhece as respostas do esquema, que levam a própria política', () => {
    expect(isViewerSchemeUrl('librevia-viewer://app/pdf-viewer/index.html')).toBe(true)
    expect(isViewerSchemeUrl('file:///app/index.html')).toBe(false)
  })
})

describe('isAllowedExternalUrl', () => {
  it.each(['https://exemplo.com', 'http://intranet.local/doc', 'mailto:alguem@exemplo.com'])(
    'aceita %s',
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
  it('permite a própria origem do servidor de desenvolvimento', () => {
    expect(isAllowedNavigation('http://localhost:5173/index.html', 'http://localhost:5173')).toBe(true)
  })

  it('recusa outra origem mesmo em desenvolvimento', () => {
    expect(isAllowedNavigation('https://exemplo.com', 'http://localhost:5173')).toBe(false)
  })

  it('permite file: apenas em produção', () => {
    expect(isAllowedNavigation('file:///app/index.html', null)).toBe(true)
    expect(isAllowedNavigation('file:///etc/passwd', 'http://localhost:5173')).toBe(false)
  })

  it('recusa URL malformada', () => {
    expect(isAllowedNavigation('%%%', null)).toBe(false)
  })
})

describe('isPdfViewerUrl', () => {
  it('reconhece a página do visualizador de PDF do Chromium', () => {
    expect(isPdfViewerUrl('chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/index.html')).toBe(true)
  })

  it('não isenta da CSP nenhuma outra origem', () => {
    // Only the viewer's own page goes without the app's CSP: anything else here would open a hole.
    expect(isPdfViewerUrl('chrome-extension://outraextensaoqualquer/index.html')).toBe(false)
    expect(isPdfViewerUrl('file:///tmp/relatorio.pdf')).toBe(false)
    expect(isPdfViewerUrl('https://mhjfbmdgcfjbbpaeojofohoefgiehjai/index.html')).toBe(false)
    expect(isPdfViewerUrl('%%%')).toBe(false)
  })
})
