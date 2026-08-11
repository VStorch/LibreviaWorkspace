import { describe, expect, it } from 'vitest'
import { normalizeLinkUrl } from './link.js'

describe('normalizeLinkUrl', () => {
  it('completa o esquema quando o usuário digita só o domínio', () => {
    expect(normalizeLinkUrl('empresa.com.br')).toBe('https://empresa.com.br/')
  })

  it('preserva um endereço completo', () => {
    expect(normalizeLinkUrl('https://intranet.local/setor/doc')).toBe('https://intranet.local/setor/doc')
  })

  it('aceita http e mailto', () => {
    expect(normalizeLinkUrl('http://intranet.local')).toBe('http://intranet.local/')
    expect(normalizeLinkUrl('mailto:setor@empresa.com.br')).toBe('mailto:setor@empresa.com.br')
  })

  it('ignora espaços em volta', () => {
    expect(normalizeLinkUrl('  https://exemplo.com  ')).toBe('https://exemplo.com/')
  })

  it.each([
    ['javascript:alert(1)', 'script disfarçado de link'],
    ['data:text/html,<script>alert(1)</script>', 'documento embutido'],
    ['file:///etc/passwd', 'arquivo local'],
    ['vbscript:msgbox', 'script legado'],
  ])('recusa %s (%s)', (url) => {
    // Um documento pode vir de qualquer lugar; link não pode virar execução.
    expect(normalizeLinkUrl(url)).toBeNull()
  })

  it('recusa endereço vazio ou sem host', () => {
    expect(normalizeLinkUrl('')).toBeNull()
    expect(normalizeLinkUrl('   ')).toBeNull()
    expect(normalizeLinkUrl('https://')).toBeNull()
  })
})
