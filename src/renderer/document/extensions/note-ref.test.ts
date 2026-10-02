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

describe('referência de nota', () => {
  it('o schema aceita a referência com o corpo dentro, e ela volta igual', () => {
    expect(doc.toJSON()).toEqual(ProseMirrorNode.fromJSON(schema, doc.toJSON()).toJSON())
    expect(noteRefsOf(doc)).toHaveLength(4)
  })

  it('numera pela ordem no texto, cada tipo por si, sem contar a marca própria', () => {
    expect(noteRefLabels(doc)).toEqual(['1', '*', 'i', '2'])
    expect(noteRefLabels(doc, { footnotePr: { numFmt: 'upperRoman' } })).toEqual(['I', '*', 'i', 'II'])
  })

  it('a colagem que repetiria uma nota do documento a leva sem nid; arrastar não', () => {
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

  it('a busca vê a nota como um bloco opaco do comprimento dela, e a contagem lê o corpo', () => {
    const paragraph = doc.child(0)
    const hidden = textWithoutDeletions(paragraph, undefined, ' ', '\u0000')
    expect(hidden.length).toBe(paragraph.content.size)
    expect(textWithoutDeletions(paragraph, ' ', ' ')).toBe('Primeira Fonte A.')
  })

  it('o texto do parágrafo para o sumário e o painel não leva o corpo da nota', () => {
    const paragraph = doc.child(1)
    expect(textBetweenWithoutNotes(paragraph, 0, paragraph.content.size)).toBe('Segunda fim')
    expect(textBetweenWithoutNotes(doc, 0, doc.content.size, '|')).toBe('Primeira|Segunda fim')
  })

  it('o HTML leva o corpo num atributo, que a colagem lê de volta', () => {
    const serializer = DOMSerializer.fromSchema(schema)
    const spec = serializer.nodes['noteRef']!(doc.child(0).child(1))
    const [tag, attrs] = spec as [string, Record<string, string>]
    expect(tag).toBe('sup')
    expect(JSON.parse(attrs['data-note-body']!)).toEqual(doc.child(0).child(1).content.toJSON())
  })

  it('na colagem, o `<sup>` da referência é nota, e não a marca de sobrescrito', () => {
    // Sem DOM no teste: a regra da nota precisa de prioridade acima da do
    // sobrescrito, que também casa `sup` e, empatada, venceria por ser marca.
    const noteRule = schema.nodes['noteRef']!.spec.parseDOM!.find((rule) => rule.tag?.startsWith('sup'))
    const supRules = (schema.marks['superscript']?.spec.parseDOM ?? []).filter(
      (rule) => 'tag' in rule && rule.tag === 'sup',
    )
    expect(noteRule?.priority ?? 50).toBeGreaterThan(
      Math.max(50, ...supRules.map((rule) => rule.priority ?? 50)),
    )
  })
})

describe('o número no começo do corpo (M11)', () => {
  it('a nota de marca própria não ganha o número: a marca já está no corpo', () => {
    const refs = noteRefsOf(doc)
    // A ordem do texto: 1, a do "*", a de fim e a 3.
    expect(refs.map(({ node }) => drawsNoteNumber(node))).toEqual([true, false, true, true])
  })
})

describe('a nota em volta de uma posição e a numeração com reinício (M11)', () => {
  const withSections = ProseMirrorNode.fromJSON(schema, {
    type: 'doc',
    content: [
      { type: 'paragraph', content: [{ type: 'text', text: 'A' }, note(null, 'Um.'), note(null, 'Dois.')] },
      { type: 'paragraph', attrs: { sectionBreak: 's1' }, content: [{ type: 'text', text: 'B' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'C' }, note(null, 'Três.')] },
    ],
  })

  it('acha a referência que contém a posição, e nenhuma fora das notas', () => {
    const [first] = noteRefsOf(withSections)
    expect(noteRefAround(withSections, first!.pos + 2)).toBe(first!.pos)
    expect(noteRefAround(withSections, 1)).toBeNull()
    expect(noteRefAround(withSections, first!.pos)).toBeNull()
  })

  it('reinicia a cada seção pelas marcas do texto', () => {
    expect(noteRefLabels(withSections, { footnotePr: { restart: 'eachSect' } })).toEqual(['1', '2', '1'])
    expect(noteRefLabels(withSections)).toEqual(['1', '2', '3'])
  })

  it('reinicia a cada página pelas folhas que a paginação dá', () => {
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
