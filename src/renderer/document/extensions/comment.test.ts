import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { EditorState, TextSelection, type Transaction } from '@tiptap/pm/state'
import { history, redo, undo } from '@tiptap/pm/history'
import type { DocumentComment } from '@services/document/model.js'
import { resolveComments } from '@services/document/comments.js'
import { buildEditorExtensions } from '../editor-extensions.js'
import { commentAnchorsOf, insertCommentAnchors, removeCommentAnchors } from './comment.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const doc = ProseMirrorNode.fromJSON(schema, {
  type: 'doc',
  content: [{ type: 'paragraph', content: [{ type: 'text', text: 'O orçamento é de doze mil.' }] }],
})

/** A biblioteca só ganha: o comentário e a resposta dele ficam nela o tempo todo. */
const library: DocumentComment[] = [
  { id: '0', author: 'Ana', date: '', paragraphs: ['Conferir.'], done: false },
  { id: '1', parentId: '0', author: 'Bruno', date: '', paragraphs: ['Ok.'], done: false },
]

const visible = (state: EditorState): string[] =>
  resolveComments(new Set(commentAnchorsOf(state.doc).keys()), library).map((comment) => comment.id)

/** Roda um comando de histórico e devolve o estado seguinte. */
function run(state: EditorState, command: typeof undo): EditorState {
  let next = state
  command(state, (tr: Transaction) => {
    next = state.apply(tr)
  })
  return next
}

describe('comentário e desfazer', () => {
  it('inserir, desfazer, refazer, excluir e desfazer de novo: o texto decide o cartão', () => {
    let state = EditorState.create({ schema, doc, plugins: [history()] })
    // "doze mil" selecionado.
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

  it('sem seleção, só a ponta de fim — o comentário de ponto', () => {
    let state = EditorState.create({ schema, doc })
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 5)))
    const tr = state.tr
    insertCommentAnchors(tr, schema, '7')
    state = state.apply(tr)
    expect(commentAnchorsOf(state.doc).get('7')).toEqual({ cid: '7', start: null, end: 5 })
  })
})
