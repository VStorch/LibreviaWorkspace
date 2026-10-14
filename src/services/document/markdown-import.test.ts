import { describe, expect, it } from 'vitest'
import { markdownToNodes } from './markdown-import.js'

describe('markdownToNodes', () => {
  it('negrito, itálico, tachado e código viram marcas', () => {
    expect(markdownToNodes('**a** *b* ~~c~~ `d`')).toEqual([
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'a', marks: [{ type: 'bold' }] },
          { type: 'text', text: ' ' },
          { type: 'text', text: 'b', marks: [{ type: 'italic' }] },
          { type: 'text', text: ' ' },
          { type: 'text', text: 'c', marks: [{ type: 'strike' }] },
          { type: 'text', text: ' ' },
          { type: 'text', text: 'd', marks: [{ type: 'code' }] },
        ],
      },
    ])
  })

  it('título guarda o nível, e parágrafos separados por linha em branco ficam separados', () => {
    expect(markdownToNodes('## Título\n\nUm\nmesmo parágrafo\n\nOutro')).toEqual([
      { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Título' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'Um mesmo parágrafo' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'Outro' }] },
    ])
  })

  it('listas com marcadores, numeradas e aninhadas', () => {
    const [bullets, numbered] = markdownToNodes('- um\n  - dentro\n- dois\n\n3. três')
    expect(bullets).toMatchObject({
      type: 'bulletList',
      content: [
        {
          type: 'listItem',
          content: [
            { type: 'paragraph', content: [{ text: 'um' }] },
            { type: 'bulletList', content: [{ type: 'listItem' }] },
          ],
        },
        { type: 'listItem', content: [{ type: 'paragraph', content: [{ text: 'dois' }] }] },
      ],
    })
    expect(numbered).toMatchObject({ type: 'orderedList', attrs: { start: 3 } })
  })

  it('citação, bloco de código e quebra de linha forçada', () => {
    expect(markdownToNodes('> citado\n\n```\nx = 1\n```\n\na  \nb')).toEqual([
      { type: 'blockquote', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'citado' }] }] },
      { type: 'codeBlock', content: [{ type: 'text', text: 'x = 1' }] },
      {
        type: 'paragraph',
        content: [{ type: 'text', text: 'a' }, { type: 'hardBreak' }, { type: 'text', text: 'b' }],
      },
    ])
  })

  it('link seguro vira link; javascript: fica só o texto', () => {
    expect(markdownToNodes('[site](https://exemplo.com) [mal](javascript:alert(1))')).toEqual([
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'site', marks: [{ type: 'link', attrs: { href: 'https://exemplo.com/' } }] },
          { type: 'text', text: ' ' },
          { type: 'text', text: 'mal' },
        ],
      },
    ])
  })

  it('HTML cru do modelo fica como texto, nunca como marcação', () => {
    expect(markdownToNodes('a <b>b</b>')).toEqual([
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'a ' },
          { type: 'text', text: '<b>' },
          { type: 'text', text: 'b' },
          { type: 'text', text: '</b>' },
        ],
      },
    ])
    expect(markdownToNodes('<script>alert(1)</script>')).toEqual([
      { type: 'paragraph', content: [{ type: 'text', text: '<script>alert(1)</script>' }] },
    ])
  })

  it('negrito dentro de negrito não repete a marca', () => {
    const [paragraph] = markdownToNodes('**a __b__**')
    expect(paragraph?.content?.[1]).toEqual({ type: 'text', text: 'b', marks: [{ type: 'bold' }] })
  })
})
