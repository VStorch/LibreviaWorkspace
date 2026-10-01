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

/** Um parágrafo de texto e pontas de comentário. */
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

describe('colar âncoras de comentário', () => {
  const copied = new Slice(Fragment.from(anchored('a ', open('0'), 'b', close('0'))), 0, 0)

  it('a cópia de um trecho comentado chega sem a âncora: a conversa já está no documento', () => {
    const doc = docWith(anchored('a ', open('0'), 'b', close('0')))
    expect(cidsIn(withoutCommentAnchors(copied, doc))).toEqual([])
  })

  it('o recortado volta com a âncora: a conversa não está mais no texto', () => {
    const doc = docWith(anchored('resto'))
    expect(cidsIn(withoutCommentAnchors(copied, doc))).toEqual(['[0', ']0'])
  })

  it('por ponta: recortada só a metade, a que saiu volta e a que ficou não se repete', () => {
    const doc = docWith(anchored(open('0'), 'antes'))
    expect(cidsIn(withoutCommentAnchors(copied, doc))).toEqual([']0'])
  })

  it('a âncora de outro documento, sem corpo aqui, sai', () => {
    const doc = docWith(anchored('resto'))
    expect(cidsIn(withoutCommentAnchors(copied, doc, false, (cid) => cid !== '0'))).toEqual([])
  })

  it('arrastar dentro do documento move, e leva a âncora', () => {
    const doc = docWith(anchored('a ', open('0'), 'b', close('0')))
    expect(withoutCommentAnchors(copied, doc, true)).toBe(copied)
  })
})

describe('próximo e anterior comentário', () => {
  // "a[0b]0 c]1 d[2e]2" — o 1 é de ponto; o 9 é resposta, sem cartão.
  const doc = docWith(
    anchored('a', open('0'), 'b', close('0'), ' c', close('1'), open('9'), close('9')),
    anchored('d', open('2'), 'e', close('2')),
  )
  const threads = new Set(['0', '1', '2'])
  const positions = commentAnchorsOf(doc)

  it('pela ordem do texto, a partir do cursor, e volta ao começo no fim', () => {
    expect(adjacentComment(doc, threads, null, 0, 1)).toBe('0')
    expect(adjacentComment(doc, threads, null, positions.get('0')!.end! + 1, 1)).toBe('1')
    expect(adjacentComment(doc, threads, null, doc.content.size, 1)).toBe('0')
    expect(adjacentComment(doc, threads, null, 0, -1)).toBe('2')
    expect(adjacentComment(doc, threads, null, positions.get('2')!.start!, -1)).toBe('1')
  })

  it('a partir da conversa em foco, com o cursor no trecho dela', () => {
    const inside = positions.get('0')!.start! + 1
    expect(adjacentComment(doc, threads, '0', inside, 1)).toBe('1')
    expect(adjacentComment(doc, threads, '0', inside, -1)).toBe('2')
    // O cursor clicado longe dela: quem manda é o cursor.
    expect(adjacentComment(doc, threads, '0', doc.content.size - 1, 1)).toBe('0')
  })

  it('sem conversa ancorada, nada', () => {
    expect(adjacentComment(docWith(anchored('nada')), threads, null, 0, 1)).toBeNull()
  })
})

describe('Backspace e Delete em volta da âncora', () => {
  it('o cursor passa por cima das âncoras, e só delas', () => {
    const doc = docWith(anchored('Fim.', close('1'), open('2')))
    const at = (pos: number) => EditorState.create({ schema, doc, selection: TextSelection.create(doc, pos) })
    // Depois das duas âncoras: o Backspace vai para depois do ".".
    expect(pastAnchors(at(7), -1)).toBe(5)
    expect(pastAnchors(at(5), 1)).toBe(7)
    expect(pastAnchors(at(4), -1)).toBeNull()
    expect(pastAnchors(at(7), 1)).toBeNull()
  })
})
