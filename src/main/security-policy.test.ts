import { describe, expect, it } from 'vitest'
import {
  SECURE_WEB_PREFERENCES,
  buildContentSecurityPolicy,
  isAllowedExternalUrl,
  isAllowedNavigation,
} from './security-policy.js'

describe('SECURE_WEB_PREFERENCES', () => {
  // Este bloco existe para que afrouxar o isolamento seja impossível de fazer
  // por acidente: quem mudar qualquer um destes valores quebra o build.
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

  it.each(['object-src', 'frame-src', 'base-uri', 'form-action'])(
    'restringe %s nos dois modos',
    (directive) => {
      expect(buildContentSecurityPolicy('production')).toContain(`${directive} 'none'`)
      expect(buildContentSecurityPolicy('development')).toContain(`${directive} 'none'`)
    },
  )
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
