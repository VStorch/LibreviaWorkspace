import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { buildEditorExtensions } from '../editor-extensions.js'
import { collectMatches } from './search-replace.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const note = (text: string, kind = 'footnote'): unknown => ({
  type: 'noteRef',
  attrs: { kind, nid: null, mark: null },
  content: [{ type: 'paragraph', content: [{ type: 'text', text }] }],
})

describe('localizar nas notas (M11)', () => {
  const doc = ProseMirrorNode.fromJSON(schema, {
    type: 'doc',
    content: [
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'ata antes' },
          note('a ata anterior'),
          { type: 'text', text: ' e ata depois' },
        ],
      },
      { type: 'paragraph', content: [{ type: 'text', text: 'fim' }, note('outra ata', 'endnote')] },
    ],
  })

  it('acha o texto das notas, na ordem do texto, sem casar através da referência', () => {
    const matches = collectMatches(doc, 'ata', false)
    expect(matches.map((match) => doc.textBetween(match.from, match.to))).toEqual([
      'ata',
      'ata',
      'ata',
      'ata',
    ])
    // A da nota fica entre a de antes e a de depois da referência.
    const reference = 1 + 'ata antes'.length
    expect(matches[0]!.from).toBeLessThan(reference)
    expect(matches[1]!.from).toBeGreaterThan(reference)
    expect(matches[2]!.from).toBeGreaterThan(reference + doc.firstChild!.child(1).nodeSize)
    expect(matches.map((match) => match.from)).toEqual(
      [...matches.map((match) => match.from)].sort((a, b) => a - b),
    )
  })

  it('a ocorrência dentro da nota aponta para dentro do corpo dela', () => {
    const inside = collectMatches(doc, 'anterior', false)
    expect(inside).toHaveLength(1)
    const $from = doc.resolve(inside[0]!.from)
    expect($from.node($from.depth - 1).type.name).toBe('noteRef')
  })
})
