import { describe, expect, it } from 'vitest'
import { plainPasteContent } from './paste.js'

describe('paste without formatting', () => {
  it('one line becomes plain text, not a new paragraph', () => {
    // Inserting a paragraph here would split the sentence in two, and pasting a city name into an
    // address is the common case.
    expect(plainPasteContent('São Paulo')).toEqual([{ type: 'text', text: 'São Paulo' }])
  })

  it('several lines become one paragraph per line', () => {
    expect(plainPasteContent('um\ndois')).toEqual([
      { type: 'paragraph', content: [{ type: 'text', text: 'um' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'dois' }] },
    ])
  })

  it('the blank line in the middle stays', () => {
    expect(plainPasteContent('um\n\ndois')).toHaveLength(3)
    expect(plainPasteContent('um\n\ndois')[1]).toEqual({ type: 'paragraph' })
  })

  it('Windows and classic Mac line breaks count as one', () => {
    expect(plainPasteContent('um\r\ndois')).toHaveLength(2)
    expect(plainPasteContent('um\rdois')).toHaveLength(2)
  })

  it('the trailing break leaves no spare empty paragraph', () => {
    // Text copied from a terminal, a spreadsheet cell or a table comes with one.
    expect(plainPasteContent('só isto\n')).toEqual([{ type: 'text', text: 'só isto' }])
  })

  it('a control character does not enter the document', () => {
    // A `\u0000` from a binary file would make the document impossible to save; `\v` is how an
    // Excel cell separates lines.
    expect(plainPasteContent('a\u0000b')).toEqual([{ type: 'text', text: 'ab' }])
    expect(plainPasteContent('a\vb')).toHaveLength(2)
  })

  it('the tab stays in the text', () => {
    // It is content: a table pasted as text loses its alignment if the tab disappears.
    expect(plainPasteContent('a\tb')).toEqual([{ type: 'text', text: 'a\tb' }])
  })

  it('empty text inserts nothing', () => {
    expect(plainPasteContent('')).toEqual([])
    expect(plainPasteContent('\n')).toEqual([])
  })
})
