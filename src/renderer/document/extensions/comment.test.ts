import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Fragment, Node as ProseMirrorNode, Slice } from '@tiptap/pm/model'
import { EditorState, TextSelection, type Transaction } from '@tiptap/pm/state'
import { history, redo, undo } from '@tiptap/pm/history'
import type { DocumentComment } from '@services/document/model.js'
import { resolveComments } from '@services/document/comments.js'
import { buildEditorExtensions } from '../editor-extensions.js'
import {
  adjacentComment,
  commentAnchorsOf,
  insertCommentAnchors,
  removeCommentAnchors,
  withoutCommentAnchors,
} from './comment.js'
import { pastAnchors } from './zero-width.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const doc = ProseMirrorNode.fromJSON(schema, {
  type: 'doc',
  content: [{ type: 'paragraph', content: [{ type: 'text', text: 'O orçamento é de doze mil.' }] }],
})

/** The library only grows: the comment and its reply stay in it the whole time. */
const library: DocumentComment[] = [
  { id: '0', author: 'Ana', date: '', paragraphs: ['Conferir.'], done: false },
  { id: '1', parentId: '0', author: 'Bruno', date: '', paragraphs: ['Ok.'], done: false },
]

const visible = (state: EditorState): string[] =>
  resolveComments(new Set(commentAnchorsOf(state.doc).keys()), library).map((comment) => comment.id)

/** Runs a history command and returns the next state. */
function run(state: EditorState, command: typeof undo): EditorState {
  let next = state
  command(state, (tr: Transaction) => {
    next = state.apply(tr)
  })
  return next
}

describe('comment and undo', () => {
  it('insert, undo, redo, delete and undo again: the text decides the card', () => {
    let state = EditorState.create({ schema, doc, plugins: [history()] })
    // "doze mil" selected.
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 18, 26)))
    expect(visible(state)).toEqual([])

    const insert = state.tr
    expect(insertCommentAnchors(insert, schema, '0')).toBe(true)
    state = state.apply(insert)
    const anchor = commentAnchorsOf(state.doc).get('0')!
    expect(state.doc.textBetween(anchor.start! + 1, anchor.end!)).toBe('doze mil')
    expect(visible(state)).toEqual(['0', '1'])

    state = run(state, undo)
    expect(visible(state)).toEqual([])
    expect(state.doc.eq(doc)).toBe(true)

    state = run(state, redo)
    expect(visible(state)).toEqual(['0', '1'])

    const remove = state.tr
    expect(removeCommentAnchors(remove, '0')).toBe(true)
    state = state.apply(remove)
    expect(visible(state)).toEqual([])
    expect(state.doc.eq(doc)).toBe(true)

    state = run(state, undo)
    expect(visible(state)).toEqual(['0', '1'])
  })

  it('with no selection, only the end anchor: the point comment', () => {
    let state = EditorState.create({ schema, doc })
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 5)))
    const tr = state.tr
    insertCommentAnchors(tr, schema, '7')
    state = state.apply(tr)
    expect(commentAnchorsOf(state.doc).get('7')).toEqual({ cid: '7', start: null, end: 5 })
  })
})

/** A text paragraph and comment ends. */
const anchored = (
  ...parts: (string | { readonly at: 'commentStart' | 'commentEnd'; readonly cid: string })[]
) =>
  ProseMirrorNode.fromJSON(schema, {
    type: 'paragraph',
    content: parts.map((part) =>
      typeof part === 'string' ? { type: 'text', text: part } : { type: part.at, attrs: { cid: part.cid } },
    ),
  })
const open = (cid: string) => ({ at: 'commentStart' as const, cid })
const close = (cid: string) => ({ at: 'commentEnd' as const, cid })
const docWith = (...paragraphs: ProseMirrorNode[]) => schema.node('doc', null, paragraphs)
const cidsIn = (slice: Slice): string[] => {
  const found: string[] = []
  slice.content.descendants((node) => {
    if (node.type.name === 'commentStart' || node.type.name === 'commentEnd')
      found.push(`${node.type.name === 'commentStart' ? '[' : ']'}${String(node.attrs['cid'])}`)
    return true
  })
  return found
}

describe('pasting comment anchors', () => {
  const copied = new Slice(Fragment.from(anchored('a ', open('0'), 'b', close('0'))), 0, 0)

  it('a copy of a commented excerpt arrives without the anchor: the thread is already in the document', () => {
    const doc = docWith(anchored('a ', open('0'), 'b', close('0')))
    expect(cidsIn(withoutCommentAnchors(copied, doc))).toEqual([])
  })

  it('the cut text comes back with the anchor: the thread is no longer in the text', () => {
    const doc = docWith(anchored('resto'))
    expect(cidsIn(withoutCommentAnchors(copied, doc))).toEqual(['[0', ']0'])
  })

  it('per anchor: with only half cut, the one that left comes back and the one that stayed is not repeated', () => {
    const doc = docWith(anchored(open('0'), 'antes'))
    expect(cidsIn(withoutCommentAnchors(copied, doc))).toEqual([']0'])
  })

  it('an anchor from another document, with no body here, is dropped', () => {
    const doc = docWith(anchored('resto'))
    expect(cidsIn(withoutCommentAnchors(copied, doc, false, (cid) => cid !== '0'))).toEqual([])
  })

  it('dragging within the document moves it, and takes the anchor', () => {
    const doc = docWith(anchored('a ', open('0'), 'b', close('0')))
    expect(withoutCommentAnchors(copied, doc, true)).toBe(copied)
  })
})

describe('next and previous comment', () => {
  // "a[0b]0 c]1 d[2e]2": 1 is a point comment; 9 is a reply, without a card.
  const doc = docWith(
    anchored('a', open('0'), 'b', close('0'), ' c', close('1'), open('9'), close('9')),
    anchored('d', open('2'), 'e', close('2')),
  )
  const threads = new Set(['0', '1', '2'])
  const positions = commentAnchorsOf(doc)

  it('in text order, from the cursor, wrapping to the start at the end', () => {
    expect(adjacentComment(doc, threads, null, 0, 1)).toBe('0')
    expect(adjacentComment(doc, threads, null, positions.get('0')!.end! + 1, 1)).toBe('1')
    expect(adjacentComment(doc, threads, null, doc.content.size, 1)).toBe('0')
    expect(adjacentComment(doc, threads, null, 0, -1)).toBe('2')
    expect(adjacentComment(doc, threads, null, positions.get('2')!.start!, -1)).toBe('1')
  })

  it('from the focused thread, with the cursor in its excerpt', () => {
    const inside = positions.get('0')!.start! + 1
    expect(adjacentComment(doc, threads, '0', inside, 1)).toBe('1')
    expect(adjacentComment(doc, threads, '0', inside, -1)).toBe('2')
    // The cursor clicked far from it: the cursor rules.
    expect(adjacentComment(doc, threads, '0', doc.content.size - 1, 1)).toBe('0')
  })

  it('with no anchored thread, nothing', () => {
    expect(adjacentComment(docWith(anchored('nada')), threads, null, 0, 1)).toBeNull()
  })
})

describe('Backspace and Delete around the anchor', () => {
  it('the cursor passes over the anchors, and only them', () => {
    const doc = docWith(anchored('Fim.', close('1'), open('2')))
    const at = (pos: number) => EditorState.create({ schema, doc, selection: TextSelection.create(doc, pos) })
    // After both anchors: Backspace goes after the ".".
    expect(pastAnchors(at(7), -1)).toBe(5)
    expect(pastAnchors(at(5), 1)).toBe(7)
    expect(pastAnchors(at(4), -1)).toBeNull()
    expect(pastAnchors(at(7), 1)).toBeNull()
  })
})

describe('comment inside a note (M11)', () => {
  const withNote = ProseMirrorNode.fromJSON(schema, {
    type: 'doc',
    content: [
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'Corpo' },
          {
            type: 'noteRef',
            attrs: { kind: 'footnote', nid: '1', mark: null },
            content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Fonte da nota.' }] }],
          },
        ],
      },
    ],
  })
  // "Fonte" inside the note: the body paragraph, the reference, the note paragraph.
  const from = 1 + 'Corpo'.length + 2
  const to = from + 'Fonte'.length

  it('the anchors go into the note, within its selection, and the panel sees them', () => {
    const state = EditorState.create({ doc: withNote })
    const tr = state.tr
    expect(insertCommentAnchors(tr, schema, '0', { from, to })).toBe(true)
    const anchor = commentAnchorsOf(tr.doc).get('0')!
    expect(tr.doc.textBetween(anchor.start! + 1, anchor.end!)).toBe('Fonte')
    expect(tr.doc.child(0).child(1).textContent).toBe('Fonte da nota.')
    expect(tr.doc.child(0).child(1).child(0).childCount).toBe(4)

    // Deleting the thread removes the ends inside the note.
    const removal = EditorState.create({ doc: tr.doc }).tr
    expect(removeCommentAnchors(removal, '0')).toBe(true)
    expect(commentAnchorsOf(removal.doc).size).toBe(0)
  })
})
