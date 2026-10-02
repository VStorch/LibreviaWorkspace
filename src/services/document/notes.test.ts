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

  it('reinicia a cada folha (eachPage) só nas de rodapé, e só com a folha conhecida', () => {
    const notes = {
      footnotePr: { restart: 'eachPage', start: 1 },
      endnotePr: { restart: 'eachPage' },
    }
    const labels = noteLabels(
      [
        { kind: NoteKind.Footnote, page: 0 },
        { kind: NoteKind.Footnote, page: 0 },
        { kind: NoteKind.Endnote, page: 0 },
        { kind: NoteKind.Footnote, page: 1 },
        { kind: NoteKind.Endnote, page: 1 },
        { kind: NoteKind.Footnote, mark: '*', page: 2 },
        { kind: NoteKind.Footnote, page: 1 },
        { kind: NoteKind.Footnote },
      ],
      notes,
    )
    expect(labels).toEqual(['1', '2', 'i', '1', 'ii', '*', '2', '3'])
  })

  it('reinicia a cada seção (eachSect), no início declarado', () => {
    const notes = { footnotePr: { restart: 'eachSect', start: 5 }, endnotePr: { restart: 'eachSect' } }
    const labels = noteLabels(
      [
        { kind: NoteKind.Footnote, section: 0, page: 0 },
        { kind: NoteKind.Footnote, section: 0, page: 1 },
        { kind: NoteKind.Endnote, section: 0 },
        { kind: NoteKind.Footnote, section: 1, page: 1 },
        { kind: NoteKind.Endnote, section: 1 },
      ],
      notes,
    )
    expect(labels).toEqual(['5', '6', 'i', '5', 'i'])
  })

  it('contínua não reinicia, mude a folha ou a seção', () => {
    const labels = noteLabels(
      [
        { kind: NoteKind.Footnote, section: 0, page: 0 },
        { kind: NoteKind.Footnote, section: 1, page: 1 },
      ],
      { footnotePr: { restart: 'continuous' } },
    )
    expect(labels).toEqual(['1', '2'])
  })
})
