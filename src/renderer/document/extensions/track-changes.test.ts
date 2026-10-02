import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { EditorState, TextSelection } from '@tiptap/pm/state'
import { history, undo } from '@tiptap/pm/history'
import { buildEditorExtensions } from '../editor-extensions.js'
import { findOccurrences } from '@services/document/search.js'
import {
  adjacentChange,
  selectChange,
  changeAt,
  revisionChangesOf,
  settleAllChanges,
  settleChangeAt,
  textWithoutDeletions,
} from './track-changes.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const ins = (author: string, rid: string): Record<string, unknown> => ({
  type: 'insertion',
  attrs: { author, rid, date: '2026-01-01T00:00:00Z' },
})
const del = (author: string, rid: string): Record<string, unknown> => ({
  type: 'deletion',
  attrs: { author, rid, date: '2026-01-01T00:00:00Z' },
})

/**
 * "Texto [inserido][ excluído] fim." — com um marcador no meio do excluído —,
 * um parágrafo cuja marca foi inserida, e uma tabela com a primeira linha excluída.
 */
const doc = ProseMirrorNode.fromJSON(schema, {
  type: 'doc',
  content: [
    {
      type: 'paragraph',
      content: [
        { type: 'text', text: 'Texto ' },
        { type: 'text', text: 'inserido', marks: [ins('Ana', '1')] },
        { type: 'text', text: ' exc', marks: [del('Bruno', '2')] },
        { type: 'bookmarkStart', attrs: { name: 'alvo', bid: '9' } },
        { type: 'text', text: 'luído', marks: [del('Bruno', '2')] },
        { type: 'bookmarkEnd', attrs: { bid: '9' } },
        { type: 'text', text: ' fim.' },
      ],
    },
    {
      type: 'paragraph',
      attrs: { markRevision: { kind: 'ins', author: 'Ana', rid: '3' } },
      content: [{ type: 'text', text: 'Partido' }],
    },
    { type: 'paragraph', content: [{ type: 'text', text: 'Seguinte' }] },
    {
      type: 'table',
      content: [
        {
          type: 'tableRow',
          attrs: { rowRevision: { kind: 'del', author: 'Bruno', rid: '4' } },
          content: [
            { type: 'tableCell', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Sai' }] }] },
          ],
        },
        {
          type: 'tableRow',
          content: [
            {
              type: 'tableCell',
              content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Fica' }] }],
            },
          ],
        },
      ],
    },
  ],
})

const firstText = (state: EditorState): string => state.doc.child(0).textContent

function positionOf(node: ProseMirrorNode, text: string): number {
  let found = -1
  node.descendants((child, pos) => {
    if (found >= 0) return false
    if (child.isText && child.text?.includes(text) === true) found = pos + child.text.indexOf(text)
    return true
  })
  return found
}

describe('controle de alterações', () => {
  it('acha cada alteração, e o marcador no meio não parte a exclusão', () => {
    const changes = revisionChangesOf(doc)
    expect(changes.map((change) => change.kind)).toEqual([
      'insertion',
      'deletion',
      'markInsertion',
      'rowDeletion',
    ])
    const deletion = changes[1]!
    expect(deletion.segments).toHaveLength(2)
    expect(doc.textBetween(deletion.from, deletion.to)).toBe(' excluído')
  })

  it('aceitar a exclusão apaga o texto e deixa o marcador', () => {
    const state = EditorState.create({ schema, doc })
    const tr = state.tr
    expect(settleChangeAt(tr, positionOf(doc, 'luído') + 1, true)).toBe(true)
    const next = state.apply(tr)
    expect(firstText(next)).toBe('Texto inserido fim.')
    const types: string[] = []
    next.doc.child(0).forEach((child) => types.push(child.type.name))
    expect(types).toContain('bookmarkStart')
    expect(types).toContain('bookmarkEnd')
  })

  it('rejeitar a exclusão devolve o texto como texto comum', () => {
    const state = EditorState.create({ schema, doc })
    const tr = state.tr
    settleChangeAt(tr, positionOf(doc, 'exc') + 1, false)
    const next = state.apply(tr)
    expect(firstText(next)).toBe('Texto inserido excluído fim.')
    expect(revisionChangesOf(next.doc).map((change) => change.kind)).not.toContain('deletion')
  })

  it('rejeitar a inserção a apaga; aceitar tira só a marca', () => {
    const state = EditorState.create({ schema, doc })
    const rejected = state.tr
    settleChangeAt(rejected, positionOf(doc, 'inserido') + 2, false)
    expect(firstText(state.apply(rejected))).toBe('Texto  excluído fim.')

    const accepted = state.tr
    settleChangeAt(accepted, positionOf(doc, 'inserido') + 2, true)
    const next = state.apply(accepted)
    expect(firstText(next)).toBe('Texto inserido excluído fim.')
    expect(revisionChangesOf(next.doc)[0]!.kind).toBe('deletion')
  })

  it('rejeitar a marca de parágrafo inserida junta com o seguinte', () => {
    const state = EditorState.create({ schema, doc })
    const tr = state.tr
    expect(changeAt(doc, positionOf(doc, 'Partido') + 2)?.kind).toBe('markInsertion')
    settleChangeAt(tr, positionOf(doc, 'Partido') + 2, false)
    const next = state.apply(tr)
    expect(next.doc.child(1).textContent).toBe('PartidoSeguinte')
    expect(next.doc.child(1).attrs['markRevision']).toBeNull()
  })

  it('aceitar tudo numa transação: um desfazer devolve tudo', () => {
    let state = EditorState.create({ schema, doc, plugins: [history()] })
    const tr = state.tr
    expect(settleAllChanges(tr, true)).toBe(true)
    state = state.apply(tr)
    expect(revisionChangesOf(state.doc)).toEqual([])
    expect(firstText(state)).toBe('Texto inserido fim.')
    expect(state.doc.child(1).textContent).toBe('Partido')
    // A linha excluída saiu; a outra fica.
    expect(state.doc.child(3).childCount).toBe(1)
    expect(state.doc.child(3).textContent).toBe('Fica')

    undo(state, (undone) => {
      state = state.apply(undone)
    })
    expect(state.doc.eq(doc)).toBe(true)
  })

  it('rejeitar tudo: o inserido sai, o excluído volta, a linha fica', () => {
    const state = EditorState.create({ schema, doc })
    const tr = state.tr
    settleAllChanges(tr, false)
    const next = state.apply(tr)
    expect(firstText(next)).toBe('Texto  excluído fim.')
    expect(next.doc.child(1).textContent).toBe('PartidoSeguinte')
    expect(next.doc.lastChild?.childCount).toBe(2)
    expect(revisionChangesOf(next.doc)).toEqual([])
  })

  it('próxima e anterior andam pela ordem do texto', () => {
    expect(adjacentChange(doc, 0, 1)?.kind).toBe('insertion')
    const insertion = adjacentChange(doc, 0, 1)!
    expect(adjacentChange(doc, insertion.to, 1)?.kind).toBe('deletion')
    const row = adjacentChange(doc, doc.content.size, -1)!
    expect(row.kind).toBe('rowDeletion')
    // Escolhida, a linha põe o cursor dentro dela; a anterior não é ela de novo.
    const inside = selectChange(EditorState.create({ schema, doc }).tr, row).selection.from
    expect(adjacentChange(doc, inside, -1)?.kind).not.toBe('rowDeletion')
    const state = EditorState.create({ schema, doc, selection: TextSelection.create(doc, 1) })
    expect(state.selection.from).toBe(1)
  })

  it('a busca e a contagem não veem o texto excluído', () => {
    const paragraph = doc.child(0)
    const hidden = textWithoutDeletions(paragraph, undefined, ' ', '\u0000')
    expect(hidden).toHaveLength(paragraph.content.size)
    expect(findOccurrences(hidden, 'luído', false)).toEqual([])
    expect(findOccurrences(hidden, 'inserido', false)).toHaveLength(1)
    expect(textWithoutDeletions(paragraph, ' ', ' ')).toBe('Texto inserido   fim.')
  })
})

describe('alterações dentro das notas (M11)', () => {
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
            content: [
              {
                type: 'paragraph',
                content: [
                  { type: 'text', text: 'Nota ' },
                  { type: 'text', text: 'nova', marks: [ins('Ana', '7')] },
                  { type: 'text', text: ' velha', marks: [del('Ana', '8')] },
                ],
              },
            ],
          },
          { type: 'text', text: ' fim' },
        ],
      },
    ],
  })

  it('acha as alterações do corpo da nota, na ordem do texto', () => {
    const changes = revisionChangesOf(withNote)
    expect(changes.map((change) => withNote.textBetween(change.from, change.to))).toEqual(['nova', ' velha'])
  })

  it('aceitar todas não deixa marca na nota', () => {
    const state = EditorState.create({ doc: withNote })
    const tr = state.tr
    expect(settleAllChanges(tr, true)).toBe(true)
    expect(revisionChangesOf(tr.doc)).toEqual([])
    expect(tr.doc.child(0).child(1).textContent).toBe('Nota nova')
  })

  it('rejeitar todas devolve a nota como era', () => {
    const tr = EditorState.create({ doc: withNote }).tr
    settleAllChanges(tr, false)
    expect(tr.doc.child(0).child(1).textContent).toBe('Nota  velha')
    expect(revisionChangesOf(tr.doc)).toEqual([])
  })

  it('aceita ou rejeita a alteração no cursor de dentro da nota, e só ela', () => {
    const inside = positionOf(withNote, 'velha') + 1
    const tr = EditorState.create({ doc: withNote }).tr
    expect(settleChangeAt(tr, inside, false)).toBe(true)
    expect(tr.doc.child(0).child(1).textContent).toBe('Nota nova velha')
    expect(revisionChangesOf(tr.doc).map((change) => tr.doc.textBetween(change.from, change.to))).toEqual([
      'nova',
    ])
  })

  it('do cursor na nota, a próxima e a anterior andam dentro da mesma nota', () => {
    const afterNova = positionOf(withNote, 'nova') + 'nova'.length
    const next = adjacentChange(withNote, afterNova, 1)!
    expect(withNote.textBetween(next.from, next.to)).toBe(' velha')
    const back = adjacentChange(withNote, next.from, -1)!
    expect(withNote.textBetween(back.from, back.to)).toBe('nova')
    expect(adjacentChange(withNote, next.to, 1)).toBeNull()
  })

  it('a próxima alteração entra na nota', () => {
    expect(
      withNote.textBetween(adjacentChange(withNote, 0, 1)!.from, adjacentChange(withNote, 0, 1)!.to),
    ).toBe('nova')
  })
})
