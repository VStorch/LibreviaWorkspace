import { describe, expect, it } from 'vitest'
import type { DocumentNode } from './model.js'
import { charactersWithoutSpaces, countParagraphs, documentStatistics } from './word-count.js'

function paragraph(text: string): DocumentNode {
  return text === '' ? { type: 'paragraph' } : { type: 'paragraph', content: [{ type: 'text', text: text }] }
}

describe('characters without spaces', () => {
  it('excludes space, tab and line break', () => {
    expect(charactersWithoutSpaces('a b\tc\nd')).toBe(4)
  })

  it('also excludes the non-breaking space that comes from Word', () => {
    // `\u00a0` appears in dates, numbers and before units. Counting it as a text character would give a
    // larger total than Word's, with nothing on screen to explain the difference.
    expect(charactersWithoutSpaces('12\u00a0kg')).toBe(4)
  })

  it('counts an accented character as one', () => {
    expect(charactersWithoutSpaces('ação')).toBe(4)
  })
})

describe('paragraphs', () => {
  it('counts a line with text and ignores an empty line', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragraph('Primeira'), paragraph(''), paragraph('Segunda')],
    }

    // An Enter pressed to make room is not content, and Word does not count it either.
    expect(countParagraphs(doc)).toBe(2)
  })

  it('counts a heading as a paragraph', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [{ type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Título' }] }],
    }

    expect(countParagraphs(doc)).toBe(1)
  })

  it('does not count twice the paragraph inside a list item or cell', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [
        {
          type: 'bulletList',
          content: [{ type: 'listItem', content: [paragraph('Item')] }],
        },
        {
          type: 'table',
          content: [
            {
              type: 'tableRow',
              content: [{ type: 'tableCell', content: [paragraph('Célula')] }],
            },
          ],
        },
      ],
    }

    expect(countParagraphs(doc)).toBe(2)
  })

  it('a section mark is not a paragraph of the text', () => {
    // In OOXML a section ends in a `w:sectPr` kept inside an empty paragraph: file structure, and
    // the corpus evidence document has six such marks in the middle of the text.
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragraph('Texto'), { type: 'paragraph', attrs: { sectionMark: true } }],
    }

    expect(countParagraphs(doc)).toBe(1)
  })
})

describe('documentStatistics', () => {
  it('measures characters and paragraphs without the deleted text and keeps what came counted', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'Olá mundo' },
            { type: 'text', text: ' velho', marks: [{ type: 'deletion' }] },
          ],
        },
        { type: 'paragraph', content: [{ type: 'text', text: 'Fim.' }] },
        { type: 'paragraph' },
      ],
    }
    expect(documentStatistics(doc, { pages: 1, words: 3, charactersWithSpaces: 13 })).toEqual({
      pages: 1,
      words: 3,
      charactersWithSpaces: 13,
      characters: 12,
      paragraphs: 2,
    })
  })
})
