import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { DOMSerializer, Node as ProseMirrorNode, Slice } from '@tiptap/pm/model'
import { buildEditorExtensions } from '../editor-extensions.js'
import { textWithoutDeletions } from './track-changes.js'
import {
  drawsNoteNumber,
  footnotePagesOf,
  noteRefAround,
  noteRefLabels,
  noteRefsOf,
  samePages,
  textBetweenWithoutNotes,
  withoutRepeatedNotes,
} from './note-ref.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const note = (nid: string | null, text: string, kind = 'footnote', mark: string | null = null): unknown => ({
  type: 'noteRef',
  attrs: { kind, nid, mark },
  content: [{ type: 'paragraph', content: [{ type: 'text', text }] }],
})

const doc = ProseMirrorNode.fromJSON(schema, {
  type: 'doc',
  content: [
    { type: 'paragraph', content: [{ type: 'text', text: 'Primeira' }, note('1', 'Fonte A.')] },
    {
      type: 'paragraph',
      content: [
        { type: 'text', text: 'Segunda' },
        note('2', 'Marca.', 'footnote', '*'),
        note('1', 'Ao fim.', 'endnote'),
        note('3', 'Fonte B.'),
        { type: 'text', text: ' fim' },
      ],
    },
  ],
})

describe('note reference', () => {
  it('the schema accepts the reference with the body inside, and it comes back the same', () => {
    expect(doc.toJSON()).toEqual(ProseMirrorNode.fromJSON(schema, doc.toJSON()).toJSON())
    expect(noteRefsOf(doc)).toHaveLength(4)
  })

  it('numbers by text order, each kind on its own, not counting custom marks', () => {
    expect(noteRefLabels(doc)).toEqual(['1', '*', 'i', '2'])
    expect(noteRefLabels(doc, { footnotePr: { numFmt: 'upperRoman' } })).toEqual(['I', '*', 'i', 'II'])
  })

  it('a paste that would duplicate a document note drops its nid; dragging does not', () => {
    const slice = new Slice(doc.child(0).content, 0, 0)
    const pasted = withoutRepeatedNotes(slice, doc)
    const ref = noteRefsOf(
      ProseMirrorNode.fromJSON(schema, {
        type: 'doc',
        content: [{ type: 'paragraph', content: pasted.content.toJSON() as unknown[] }],
      }),
    )[0]!
    expect(ref.node.attrs['nid']).toBeNull()
    expect(ref.node.textContent).toBe('Fonte A.')
    expect(withoutRepeatedNotes(slice, doc, true)).toBe(slice)
  })

  it('search sees the note as an opaque block of its length, and counting reads the body', () => {
    const paragraph = doc.child(0)
    const hidden = textWithoutDeletions(paragraph, undefined, ' ', '\u0000')
    expect(hidden.length).toBe(paragraph.content.size)
    expect(textWithoutDeletions(paragraph, ' ', ' ')).toBe('Primeira Fonte A.')
  })

  it('the paragraph text for the outline and the panel does not include the note body', () => {
    const paragraph = doc.child(1)
    expect(textBetweenWithoutNotes(paragraph, 0, paragraph.content.size)).toBe('Segunda fim')
    expect(textBetweenWithoutNotes(doc, 0, doc.content.size, '|')).toBe('Primeira|Segunda fim')
  })

  it('the HTML carries the body in an attribute, which paste reads back', () => {
    const serializer = DOMSerializer.fromSchema(schema)
    const spec = serializer.nodes['noteRef']!(doc.child(0).child(1))
    const [tag, attrs] = spec as [string, Record<string, string>]
    expect(tag).toBe('sup')
    expect(JSON.parse(attrs['data-note-body']!)).toEqual(doc.child(0).child(1).content.toJSON())
  })

  it("on paste, the reference's `<sup>` is a note, not the superscript mark", () => {
    // No DOM in the test: the note rule needs a higher priority than superscript, which also
    // matches `sup` and, on a tie, would win for being a mark.
    const noteRule = schema.nodes['noteRef']!.spec.parseDOM!.find((rule) => rule.tag?.startsWith('sup'))
    const supRules = (schema.marks['superscript']?.spec.parseDOM ?? []).filter(
      (rule) => 'tag' in rule && rule.tag === 'sup',
    )
    expect(noteRule?.priority ?? 50).toBeGreaterThan(
      Math.max(50, ...supRules.map((rule) => rule.priority ?? 50)),
    )
  })
})

describe('the number at the start of the body (M11)', () => {
  it('a note with a custom mark gets no number: the mark is already in the body', () => {
    const refs = noteRefsOf(doc)
    // Text order: 1, the "*" one, the endnote and 3.
    expect(refs.map(({ node }) => drawsNoteNumber(node))).toEqual([true, false, true, true])
  })
})

describe('the note around a position and numbering with restart (M11)', () => {
  const withSections = ProseMirrorNode.fromJSON(schema, {
    type: 'doc',
    content: [
      { type: 'paragraph', content: [{ type: 'text', text: 'A' }, note(null, 'Um.'), note(null, 'Dois.')] },
      { type: 'paragraph', attrs: { sectionBreak: 's1' }, content: [{ type: 'text', text: 'B' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'C' }, note(null, 'Três.')] },
    ],
  })

  it('finds the reference containing the position, and none outside the notes', () => {
    const [first] = noteRefsOf(withSections)
    expect(noteRefAround(withSections, first!.pos + 2)).toBe(first!.pos)
    expect(noteRefAround(withSections, 1)).toBeNull()
    expect(noteRefAround(withSections, first!.pos)).toBeNull()
  })

  it('restarts at each section by the text marks', () => {
    expect(noteRefLabels(withSections, { footnotePr: { restart: 'eachSect' } })).toEqual(['1', '2', '1'])
    expect(noteRefLabels(withSections)).toEqual(['1', '2', '3'])
  })

  it('restarts at each page by the sheets pagination gives', () => {
    const areas = [
      { sheet: 0, kind: 'footnote', items: [{ index: 0, fromLine: 0 }] },
      {
        sheet: 1,
        kind: 'footnote',
        items: [
          { index: 0, fromLine: 3 },
          { index: 1, fromLine: 0 },
          { index: 2, fromLine: 0 },
        ],
      },
    ]
    const pages = footnotePagesOf(areas)
    expect(pages).toEqual([0, 1, 1])
    expect(noteRefLabels(withSections, { footnotePr: { restart: 'eachPage' } }, pages)).toEqual([
      '1',
      '1',
      '2',
    ])
    expect(samePages(pages, [0, 1, 1])).toBe(true)
    expect(samePages(pages, [0, 1])).toBe(false)
  })
})
