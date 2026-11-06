import { describe, expect, it } from 'vitest'
import { normalizeLinkUrl, withoutUnsafeLinks } from './link.js'

describe('normalizeLinkUrl', () => {
  it('completes the scheme when the user types only the domain', () => {
    expect(normalizeLinkUrl('empresa.com.br')).toBe('https://empresa.com.br/')
  })

  it('preserves a complete address', () => {
    expect(normalizeLinkUrl('https://intranet.local/setor/doc')).toBe('https://intranet.local/setor/doc')
  })

  it('accepts http and mailto', () => {
    expect(normalizeLinkUrl('http://intranet.local')).toBe('http://intranet.local/')
    expect(normalizeLinkUrl('mailto:setor@empresa.com.br')).toBe('mailto:setor@empresa.com.br')
  })

  it('ignores surrounding spaces', () => {
    expect(normalizeLinkUrl('  https://exemplo.com  ')).toBe('https://exemplo.com/')
  })

  it.each([
    ['javascript:alert(1)', 'script disguised as a link'],
    ['data:text/html,<script>alert(1)</script>', 'embedded document'],
    ['file:///etc/passwd', 'local file'],
    ['vbscript:msgbox', 'legacy script'],
  ])('refuses %s (%s)', (url) => {
    // A document can come from anywhere; a link must not become execution.
    expect(normalizeLinkUrl(url)).toBeNull()
  })

  it('rejects an empty address or one without a host', () => {
    expect(normalizeLinkUrl('')).toBeNull()
    expect(normalizeLinkUrl('   ')).toBeNull()
    expect(normalizeLinkUrl('https://')).toBeNull()
  })
})

describe('withoutUnsafeLinks', () => {
  const link = (href: string) => ({ type: 'link', attrs: { href } })

  it('unlinks a disguised script and keeps the text and the other marks', () => {
    const doc = {
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'perigo', marks: [link('javascript:alert(1)'), { type: 'bold' }] },
            { type: 'text', text: 'só link', marks: [link('file:///etc/passwd')] },
          ],
        },
      ],
    }

    expect(withoutUnsafeLinks(doc)).toEqual({
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'perigo', marks: [{ type: 'bold' }] },
            { type: 'text', text: 'só link' },
          ],
        },
      ],
    })
  })

  it('keeps web, mail and internal addresses', () => {
    const marks = [link('https://exemplo.com'), link('mailto:a@b.com'), link('#Intro')]
    expect(withoutUnsafeLinks({ type: 'text', text: 'x', marks })).toEqual({ type: 'text', text: 'x', marks })
  })

  it('passes through what is not a node', () => {
    expect(withoutUnsafeLinks(null)).toBeNull()
    expect(withoutUnsafeLinks('texto')).toBe('texto')
  })
})
