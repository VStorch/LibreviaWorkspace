import { describe, expect, it } from 'vitest'
import { NoteKind, noteLabel, noteLabels } from './notes.js'

describe('note numbering', () => {
  it('counts footnotes and endnotes separately, as Word does', () => {
    const labels = noteLabels([
      { kind: NoteKind.Footnote },
      { kind: NoteKind.Endnote },
      { kind: NoteKind.Footnote },
      { kind: NoteKind.Endnote },
    ])
    expect(labels).toEqual(['1', 'i', '2', 'ii'])
  })

  it('follows the format and start the document declares', () => {
    const notes = { footnotePr: { numFmt: 'upperLetter', start: 3 }, endnotePr: { numFmt: 'decimal' } }
    expect(noteLabels([{ kind: NoteKind.Footnote }, { kind: NoteKind.Footnote }], notes)).toEqual(['C', 'D'])
    expect(noteLabel(NoteKind.Endnote, 0, notes)).toBe('1')
    expect(noteLabel(NoteKind.Footnote, 0, { footnotePr: { numFmt: 'lowerRoman', start: 4 } })).toBe('iv')
  })

  it('a custom mark appears in place of the number and does not count', () => {
    const labels = noteLabels([
      { kind: NoteKind.Footnote },
      { kind: NoteKind.Footnote, mark: '*' },
      { kind: NoteKind.Footnote },
    ])
    expect(labels).toEqual(['1', '*', '2'])
  })

  it('chicago uses the symbols and then doubles them', () => {
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

  it('a format without a number comes out as decimal', () => {
    expect(noteLabel(NoteKind.Footnote, 1, { footnotePr: { numFmt: 'none' } })).toBe('2')
  })

  it('restarts on each sheet (eachPage) only for footnotes, and only with the sheet known', () => {
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

  it('restarts on each section (eachSect), at the declared start', () => {
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

  it('continuous does not restart, whether the sheet or the section changes', () => {
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
