import { describe, expect, it } from 'vitest'
import { documentToPlainText, hasRichFormatting, plainTextToDocument } from './plain-text.js'
import type { DocumentNode } from './model.js'

describe('plainTextToDocument', () => {
  it('turns each line into a paragraph', () => {
    const doc = plainTextToDocument('primeira\nsegunda')
    expect(doc.content).toHaveLength(2)
    expect(doc.content?.[0]).toEqual({ type: 'paragraph', content: [{ type: 'text', text: 'primeira' }] })
  })

  it('preserves blank lines', () => {
    expect(plainTextToDocument('a\n\nb').content).toHaveLength(3)
  })

  it.each([
    ['\n', 'unix'],
    ['\r\n', 'windows'],
    ['\r', 'classic mac'],
  ])('accepts line break %j (%s)', (breakChar) => {
    expect(plainTextToDocument(`a${breakChar}b`).content).toHaveLength(2)
  })

  it('never produces an empty document', () => {
    // ProseMirror refuses a doc without content; an empty file must not bring the editor down.
    expect(plainTextToDocument('').content).toEqual([{ type: 'paragraph' }])
  })
})

describe('documentToPlainText', () => {
  it('joins blocks with a line break', () => {
    const doc = plainTextToDocument('uma\noutra')
    expect(documentToPlainText(doc)).toBe('uma\noutra')
  })

  it('survives a round trip through plain text', () => {
    const original = 'linha 1\n\nlinha 3\nlinha 4'
    expect(documentToPlainText(plainTextToDocument(original))).toBe(original)
  })

  it('extracts the text of headings and lists', () => {
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

  it('ignores formatting marks and keeps only the text', () => {
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
  it('reports no formatting in a plain text document', () => {
    expect(hasRichFormatting(plainTextToDocument('só texto\nem duas linhas'))).toBe(false)
  })

  it.each([
    ['bold mark', { type: 'paragraph', content: [{ type: 'text', marks: [{ type: 'bold' }], text: 'x' }] }],
    ['heading', { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'x' }] }],
    ['alignment', { type: 'paragraph', attrs: { textAlign: 'center' } }],
    ['image', { type: 'image', attrs: { src: 'data:image/png;base64,AAA' } }],
    ['table', { type: 'table', content: [] }],
  ])('flags formatting in %s', (_label, node) => {
    // Pessimistic on purpose: this triggers the warning before saving as .txt.
    expect(hasRichFormatting({ type: 'doc', content: [node as DocumentNode] })).toBe(true)
  })
})
