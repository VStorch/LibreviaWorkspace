import { describe, expect, it } from 'vitest'
import { plainPasteContent } from './paste.js'

describe('colar sem formatação', () => {
  it('uma linha vira texto puro, e não parágrafo novo', () => {
    // Inserting a paragraph here would split the sentence in two, and pasting a city name into an
    // address is the common case.
    expect(plainPasteContent('São Paulo')).toEqual([{ type: 'text', text: 'São Paulo' }])
  })

  it('várias linhas viram um parágrafo por linha', () => {
    expect(plainPasteContent('um\ndois')).toEqual([
      { type: 'paragraph', content: [{ type: 'text', text: 'um' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'dois' }] },
    ])
  })

  it('a linha em branco do meio continua lá', () => {
    expect(plainPasteContent('um\n\ndois')).toHaveLength(3)
    expect(plainPasteContent('um\n\ndois')[1]).toEqual({ type: 'paragraph' })
  })

  it('quebra de linha do Windows e do Mac clássico contam como uma', () => {
    expect(plainPasteContent('um\r\ndois')).toHaveLength(2)
    expect(plainPasteContent('um\rdois')).toHaveLength(2)
  })

  it('a quebra do fim não deixa parágrafo vazio sobrando', () => {
    // Text copied from a terminal, a spreadsheet cell or a table comes with one.
    expect(plainPasteContent('só isto\n')).toEqual([{ type: 'text', text: 'só isto' }])
  })

  it('caractere de controle não entra no documento', () => {
    // A `\u0000` from a binary file would make the document impossible to save; `\v` is how an
    // Excel cell separates lines.
    expect(plainPasteContent('a\u0000b')).toEqual([{ type: 'text', text: 'ab' }])
    expect(plainPasteContent('a\vb')).toHaveLength(2)
  })

  it('tabulação continua no texto', () => {
    // It is content: a table pasted as text loses its alignment if the tab disappears.
    expect(plainPasteContent('a\tb')).toEqual([{ type: 'text', text: 'a\tb' }])
  })

  it('texto vazio não insere nada', () => {
    expect(plainPasteContent('')).toEqual([])
    expect(plainPasteContent('\n')).toEqual([])
  })
})
