import { beforeEach, describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { EditorState, type Transaction } from '@tiptap/pm/state'
import type { Editor } from '@tiptap/react'
import type { AiCommentProposal, AiProposal } from '@services/ai/proposals.js'
import { usePreferences } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import { buildTrackedEdits } from './ai-apply.js'
import { chatParagraphs, locateProposal } from './ai-proposals.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { applyProposals } from './proposal-commands.js'

const schema = getSchema(buildEditorExtensions(() => {}))
const NOW = new Date('2026-10-01T12:00:00Z')
const MODEL = 'modelo-falso'

const text = (value: string, ...marks: Record<string, unknown>[]): Record<string, unknown> =>
  marks.length === 0 ? { type: 'text', text: value } : { type: 'text', text: value, marks }
const paragraph = (...content: Record<string, unknown>[]): Record<string, unknown> => ({
  type: 'paragraph',
  ...(content.length === 0 ? {} : { content }),
})
const deletion = { type: 'deletion', attrs: { author: 'Ana', date: '2026-01-01T00:00:00Z', rid: '1' } }
const cell = (value: string): Record<string, unknown> => ({
  type: 'tableCell',
  content: [paragraph(text(value))],
})
const bullets = (...items: string[]): Record<string, unknown> => ({
  type: 'bulletList',
  content: items.map((item) => ({ type: 'listItem', content: [paragraph(text(item))] })),
})

function docOf(...blocks: Record<string, unknown>[]): ProseMirrorNode {
  return ProseMirrorNode.fromJSON(schema, { type: 'doc', content: blocks })
}

/** Inserted text in braces, deleted in brackets, as the person sees them in the margin. */
function shown(doc: ProseMirrorNode): string[] {
  const blocks: string[] = []
  doc.descendants((node) => {
    if (!node.isTextblock) return true
    let line = ''
    node.forEach((child) => {
      const names = child.marks.map((mark) => mark.type.name)
      const value = child.text ?? (child.type.name.startsWith('comment') ? '|' : '')
      if (names.includes('deletion')) line += `[${value}]`
      else if (names.includes('insertion')) line += `{${value}}`
      else line += value
    })
    blocks.push(line)
    return false
  })
  return blocks
}

function authorsOf(doc: ProseMirrorNode): string[] {
  const authors = new Set<string>()
  doc.descendants((node) => {
    for (const mark of node.marks) {
      if (mark.type.name === 'insertion') authors.add(String(mark.attrs['author']))
    }
  })
  return [...authors]
}

/** As main makes them: with the text of the paragraph the model read. */
const edit = (
  doc: ProseMirrorNode,
  id: string,
  paragraphNumber: number,
  find: string,
  replace: string,
): AiProposal => ({
  id,
  kind: 'edit',
  paragraph: paragraphNumber,
  paragraphText: chatParagraphs(doc)[paragraphNumber - 1] ?? '',
  find,
  replace,
})

const comment = (
  doc: ProseMirrorNode,
  id: string,
  paragraphNumber: number,
  find: string,
): AiCommentProposal => ({
  id,
  kind: 'comment',
  paragraph: paragraphNumber,
  paragraphText: chatParagraphs(doc)[paragraphNumber - 1] ?? '',
  find,
  comment: `Comentário ${id}`,
})

interface FakeEditor {
  readonly editor: Editor
  readonly transactions: Transaction[]
  state(): EditorState
}

function fakeEditor(doc: ProseMirrorNode): FakeEditor {
  let state = EditorState.create({ doc })
  const transactions: Transaction[] = []
  const editor = {
    get state() {
      return state
    },
    view: {
      dispatch(tr: Transaction) {
        transactions.push(tr)
        state = state.apply(tr)
      },
      focus: () => undefined,
    },
  } as unknown as Editor
  return { editor, transactions, state: () => state }
}

describe('numbered paragraphs for the chat', () => {
  it('one item per text block, with cells and list items, without deletions', () => {
    const doc = docOf(
      paragraph(text('O gato '), text('não ', deletion), text('dorme.')),
      { type: 'table', content: [{ type: 'tableRow', content: [cell('um'), cell('dois')] }] },
      bullets('item a', 'item b'),
      paragraph(),
    )
    expect(chatParagraphs(doc)).toEqual(['O gato dorme.', 'um', 'dois', 'item a', 'item b', ''])
  })
})

describe('locating a proposal', () => {
  it('maps the excerpt of the text without deletions to document positions', () => {
    const doc = docOf(
      paragraph(text('Primeiro.')),
      paragraph(text('O gato '), text('não ', deletion), text('dorme.')),
    )
    const range = locateProposal(doc, edit(doc, '1', 2, 'gato dorme', 'x'))
    expect(range).not.toBeNull()
    expect(doc.textBetween(range!.from, range!.to)).toBe('gato não dorme')
  })

  it('missing or repeated text, or a paragraph that does not exist: nothing to apply', () => {
    const doc = docOf(paragraph(text('a gato a')))
    expect(locateProposal(doc, edit(doc, '1', 1, 'cão', 'x'))).toBeNull()
    expect(locateProposal(doc, edit(doc, '1', 1, 'a', 'x'))).toBeNull()
    expect(locateProposal(doc, edit(doc, '1', 2, 'gato', 'x'))).toBeNull()
  })
})

describe('applying proposals', () => {
  beforeEach(() => {
    useWorkspace.setState({ readOnly: false, beforeComments: false, comments: [] })
    usePreferences.setState({ preferences: { ...usePreferences.getState().preferences, commentsPane: true } })
  })

  it('a change goes in as a tracked revision with the model as author', () => {
    const state = EditorState.create({ doc: docOf(paragraph(text('O gato dorme.'))) })
    const range = locateProposal(state.doc, edit(state.doc, '1', 1, 'dorme', 'cochila'))!
    const next = state.apply(buildTrackedEdits(state, [{ range, text: 'cochila' }], MODEL, NOW)!)
    expect(shown(next.doc)).toEqual(['O gato [dorme]{cochila}.'])
    expect(authorsOf(next.doc)).toEqual([MODEL])
  })

  it('Apply all: the changes go in a single transaction, back to front', () => {
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.')), paragraph(text('O cão late.'))))
    const outcome = applyProposals(
      fake.editor,
      [
        edit(fake.state().doc, '1', 1, 'gato', 'felino'),
        edit(fake.state().doc, '2', 1, 'dorme', 'cochila'),
        edit(fake.state().doc, '3', 2, 'late', ''),
      ],
      MODEL,
    )
    expect(outcome).toEqual({ applied: ['1', '2', '3'], changed: [] })
    expect(fake.transactions).toHaveLength(1)
    expect(shown(fake.state().doc)).toEqual(['O [gato]{felino} [dorme]{cochila}.', 'O cão [late].'])
  })

  it('text changed since the proposal is not applied and is flagged', () => {
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.'))))
    const outcome = applyProposals(
      fake.editor,
      [edit(fake.state().doc, '1', 1, 'cachorro', 'cão'), edit(fake.state().doc, '2', 1, 'gato', 'x')],
      MODEL,
    )
    expect(outcome).toEqual({ applied: ['2'], changed: ['1'] })
  })

  it('a new paragraph above renumbers: the proposal does not land on the paragraph that took the number', () => {
    const before = docOf(paragraph(text('O gato dorme.')), paragraph(text('O gato late.')))
    const proposal = edit(before, '1', 2, 'gato', 'cão')
    const fake = fakeEditor(
      docOf(paragraph(text('Novo.')), paragraph(text('O gato dorme.')), paragraph(text('O gato late.'))),
    )
    expect(applyProposals(fake.editor, [proposal], MODEL)).toEqual({ applied: [], changed: ['1'] })
    expect(fake.transactions).toHaveLength(0)
  })

  it('Apply all with comments and changes: one transaction, one undo', () => {
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.')), paragraph(text('O cão late.'))))
    const doc = fake.state().doc
    const outcome = applyProposals(
      fake.editor,
      [comment(doc, '1', 1, 'gato'), edit(doc, '2', 1, 'dorme', 'cochila'), comment(doc, '3', 2, 'late')],
      MODEL,
    )
    expect(outcome).toEqual({ applied: ['1', '2', '3'], changed: [] })
    expect(fake.transactions).toHaveLength(1)
    expect(shown(fake.state().doc)).toEqual(['O |gato| [dorme]{cochila}.', 'O cão |late|.'])
    expect(useWorkspace.getState().comments).toHaveLength(2)
  })

  it('overlapping proposals: the second is flagged as changed text', () => {
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.'))))
    const outcome = applyProposals(
      fake.editor,
      [edit(fake.state().doc, '1', 1, 'gato dorme', 'x'), edit(fake.state().doc, '2', 1, 'dorme', 'y')],
      MODEL,
    )
    expect(outcome).toEqual({ applied: ['1'], changed: ['2'] })
  })

  it('a comment goes in at both ends of the excerpt, with the model as author', () => {
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.'))))
    const proposal: AiProposal = { ...comment(fake.state().doc, '1', 1, 'gato'), comment: 'Qual gato?' }
    const outcome = applyProposals(fake.editor, [proposal], MODEL)
    expect(outcome).toEqual({ applied: ['1'], changed: [] })
    expect(shown(fake.state().doc)).toEqual(['O |gato| dorme.'])
    expect(useWorkspace.getState().comments).toMatchObject([
      { author: MODEL, paragraphs: ['Qual gato?'], done: false },
    ])
  })

  it('read-only document: nothing is applied', () => {
    useWorkspace.setState({ readOnly: true })
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.'))))
    const proposal = comment(fake.state().doc, '1', 1, 'gato')
    expect(
      applyProposals(fake.editor, [proposal, edit(fake.state().doc, '2', 1, 'dorme', 'x')], MODEL).applied,
    ).toEqual([])
    expect(fake.transactions).toHaveLength(0)
  })
})
