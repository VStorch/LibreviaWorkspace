import { describe, expect, it } from 'vitest'
import { documentToPlainText, hasRichFormatting, plainTextToDocument } from './plain-text.js'
import type { DocumentNode } from './model.js'

describe('plainTextToDocument', () => {
  it('transforma cada linha num parágrafo', () => {
    const doc = plainTextToDocument('primeira\nsegunda')
    expect(doc.content).toHaveLength(2)
    expect(doc.content?.[0]).toEqual({ type: 'paragraph', content: [{ type: 'text', text: 'primeira' }] })
  })

  it('preserva linhas em branco', () => {
    expect(plainTextToDocument('a\n\nb').content).toHaveLength(3)
  })

  it.each([
    ['\n', 'unix'],
    ['\r\n', 'windows'],
    ['\r', 'mac clássico'],
  ])('aceita quebra de linha %j (%s)', (breakChar) => {
    expect(plainTextToDocument(`a${breakChar}b`).content).toHaveLength(2)
  })

  it('nunca produz documento vazio', () => {
    // O ProseMirror recusa um doc sem conteúdo; um arquivo vazio não pode
    // derrubar o editor.
    expect(plainTextToDocument('').content).toEqual([{ type: 'paragraph' }])
  })
})

describe('documentToPlainText', () => {
  it('junta blocos com quebra de linha', () => {
    const doc = plainTextToDocument('uma\noutra')
    expect(documentToPlainText(doc)).toBe('uma\noutra')
  })

  it('sobrevive à ida e volta pelo texto puro', () => {
    const original = 'linha 1\n\nlinha 3\nlinha 4'
    expect(documentToPlainText(plainTextToDocument(original))).toBe(original)
  })

  it('extrai o texto de títulos e listas', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [
        { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Título' }] },
        {
          type: 'bulletList',
          content: [
            { type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'item' }] }] },
          ],
        },
      ],
    }

    expect(documentToPlainText(doc)).toContain('Título')
    expect(documentToPlainText(doc)).toContain('item')
  })

  it('ignora marcas de formatação e mantém só o texto', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [
            { type: 'text', marks: [{ type: 'bold' }], text: 'forte' },
            { type: 'text', text: ' normal' },
          ],
        },
      ],
    }

    expect(documentToPlainText(doc)).toBe('forte normal')
  })
})

describe('hasRichFormatting', () => {
  it('não acusa formatação num documento de texto puro', () => {
    expect(hasRichFormatting(plainTextToDocument('só texto\nem duas linhas'))).toBe(false)
  })

  it.each([
    [
      'marca de negrito',
      { type: 'paragraph', content: [{ type: 'text', marks: [{ type: 'bold' }], text: 'x' }] },
    ],
    ['título', { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'x' }] }],
    ['alinhamento', { type: 'paragraph', attrs: { textAlign: 'center' } }],
    ['imagem', { type: 'image', attrs: { src: 'data:image/png;base64,AAA' } }],
    ['tabela', { type: 'table', content: [] }],
  ])('acusa formatação em %s', (_label, node) => {
    // Pessimista de propósito: é o que dispara o aviso antes de salvar em .txt.
    expect(hasRichFormatting({ type: 'doc', content: [node as DocumentNode] })).toBe(true)
  })
})
