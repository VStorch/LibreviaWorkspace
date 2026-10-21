import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { EditorState, TextSelection } from '@tiptap/pm/state'
import { RevisionView } from '@shared/types.js'
import { buildEditorExtensions } from '../editor-extensions.js'
import {
  deleteVisible,
  hiddenRuns,
  isHiddenBlock,
  visiblePosition,
  visibleSelection,
} from './revision-view.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const mark = (type: 'insertion' | 'deletion'): Record<string, unknown> => ({ type, attrs: { author: 'Ana' } })
const text = (value: string, ...marks: Record<string, unknown>[]): Record<string, unknown> =>
  marks.length === 0 ? { type: 'text', text: value } : { type: 'text', text: value, marks }

function docOf(...blocks: Record<string, unknown>[]): ProseMirrorNode {
  return ProseMirrorNode.fromJSON(schema, { type: 'doc', content: blocks })
}

// "ab" + [cd deleted] + {ef inserted} + "gh": positions 1..9 in the first paragraph.
const doc = docOf(
  {
    type: 'paragraph',
    content: [text('ab'), text('cd', mark('deletion')), text('ef', mark('insertion')), text('gh')],
  },
  {
    type: 'paragraph',
    attrs: { markRevision: { kind: 'del', author: 'Ana' } },
    content: [text('tudo fora', mark('deletion'))],
  },
  { type: 'paragraph', content: [text('fim')] },
)

describe('the cursor outside what the mode hides', () => {
  it('each mode hides its own excerpt', () => {
    const block = doc.firstChild!
    expect(hiddenRuns(block, RevisionView.All)).toEqual([])
    expect(hiddenRuns(block, RevisionView.None)).toEqual([{ from: 2, to: 4 }])
    expect(hiddenRuns(block, RevisionView.Original)).toEqual([{ from: 4, to: 6 }])
  })

  it('inside hidden deleted text, the cursor goes to the edge in the direction it was moving', () => {
    expect(visiblePosition(doc, 4, RevisionView.Simple, 1)).toBe(5)
    expect(visiblePosition(doc, 4, RevisionView.Simple, -1)).toBe(3)
    // At the edges it is already outside; and with all markup nothing is hidden.
    expect(visiblePosition(doc, 3, RevisionView.Simple, 1)).toBe(3)
    expect(visiblePosition(doc, 4, RevisionView.All, 1)).toBe(4)
  })

  it('a fully deleted paragraph disappears, and the cursor moves to the neighbor', () => {
    const second = doc.child(1)
    expect(isHiddenBlock(second, RevisionView.None)).toBe(true)
    expect(isHiddenBlock(second, RevisionView.Original)).toBe(false)
    const inside = doc.firstChild!.nodeSize + 3
    const after = visiblePosition(doc, inside, RevisionView.None, 1)
    expect(doc.resolve(after).parent.textContent).toBe('fim')
    const before = visiblePosition(doc, inside, RevisionView.None, -1)
    expect(doc.resolve(before).parent.textContent).toBe('abcdefgh')
  })

  it('a selection already outside stays as it is', () => {
    expect(visibleSelection(doc, TextSelection.create(doc, 2), RevisionView.None, 1)).toBeNull()
    const moved = visibleSelection(doc, TextSelection.create(doc, 6), RevisionView.Original, -1)
    expect(moved?.head).toBe(5)
  })
})

describe('deleting what is visible', () => {
  it('a selection that spans hidden deleted text removes only the rest', () => {
    const tr = EditorState.create({ doc }).tr
    // "ab[cd]ef gh" from 1 to 9: without markup, "cd" is not visible and stays.
    deleteVisible(tr, 1, 9, RevisionView.None)
    expect(tr.doc.child(0).textContent).toBe('cd')
    // With everything visible, everything goes.
    const all = deleteVisible(EditorState.create({ doc }).tr, 1, 9, RevisionView.All)
    expect(all.doc.child(0).textContent).toBe('')
  })
})

describe('the cursor outside many hidden paragraphs', () => {
  it('passes more than sixteen in a row', () => {
    const hidden = Array.from({ length: 20 }, () => ({
      type: 'paragraph',
      attrs: { markRevision: { kind: 'del', author: 'Ana' } },
      content: [text('fora', mark('deletion'))],
    }))
    const long = docOf({ type: 'paragraph', content: [text('início')] }, ...hidden, {
      type: 'paragraph',
      content: [text('fim')],
    })
    const inside = 1 + 8 + 2
    const landed = visiblePosition(long, inside, RevisionView.None, 1)
    expect(long.resolve(landed).parent.textContent).toBe('fim')
  })
})
