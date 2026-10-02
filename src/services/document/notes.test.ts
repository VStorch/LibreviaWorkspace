import { describe, expect, it } from 'vitest'
import { NoteKind, noteLabel, noteLabels } from './notes.js'

describe('numeração das notas', () => {
  it('conta as de rodapé e as de fim separadas, no padrão do Word', () => {
    const labels = noteLabels([
      { kind: NoteKind.Footnote },
      { kind: NoteKind.Endnote },
      { kind: NoteKind.Footnote },
      { kind: NoteKind.Endnote },
    ])
    expect(labels).toEqual(['1', 'i', '2', 'ii'])
  })

  it('segue o formato e o início que o documento declara', () => {
    const notes = { footnotePr: { numFmt: 'upperLetter', start: 3 }, endnotePr: { numFmt: 'decimal' } }
    expect(noteLabels([{ kind: NoteKind.Footnote }, { kind: NoteKind.Footnote }], notes)).toEqual(['C', 'D'])
    expect(noteLabel(NoteKind.Endnote, 0, notes)).toBe('1')
    expect(noteLabel(NoteKind.Footnote, 0, { footnotePr: { numFmt: 'lowerRoman', start: 4 } })).toBe('iv')
  })

  it('a marca própria aparece no lugar do número e não entra na conta', () => {
    const labels = noteLabels([
      { kind: NoteKind.Footnote },
      { kind: NoteKind.Footnote, mark: '*' },
      { kind: NoteKind.Footnote },
    ])
    expect(labels).toEqual(['1', '*', '2'])
  })

  it('chicago usa os símbolos e depois os dobra', () => {
    const notes = { footnotePr: { numFmt: 'chicago' } }
    expect([0, 1, 2, 3, 4, 5].map((ordinal) => noteLabel(NoteKind.Footnote, ordinal, notes))).toEqual([
      '*',
      '†',
      '‡',
      '§',
      '**',
      '††',
    ])
  })

  it('formato sem número sai em decimal', () => {
    expect(noteLabel(NoteKind.Footnote, 1, { footnotePr: { numFmt: 'none' } })).toBe('2')
  })
})
