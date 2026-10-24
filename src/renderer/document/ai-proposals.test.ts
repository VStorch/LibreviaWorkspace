import { beforeEach, describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { EditorState, type Transaction } from '@tiptap/pm/state'
import type { Editor } from '@tiptap/react'
import type {
  AiCommentProposal,
  AiEditProposal,
  AiInsertProposal,
  AiProposal,
} from '@services/ai/proposals.js'
import { usePreferences } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import { settleAllChanges } from './extensions/track-changes.js'
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
): AiEditProposal => ({
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

const insertion = (doc: ProseMirrorNode, id: string, after: number, content: string): AiInsertProposal => ({
  id,
  kind: 'insert',
  paragraph: after,
  paragraphText: after === 0 ? '' : (chatParagraphs(doc)[after - 1] ?? ''),
  content,
})

/** Block type, then its text as `shown` writes it. */
function outline(doc: ProseMirrorNode): string[] {
  const blocks: string[] = []
  doc.forEach((node) => {
    const level = node.type.name === 'heading' ? String(node.attrs['level']) : ''
    blocks.push(
      `${node.type.name}${level}:${shown(docOf(node.toJSON() as Record<string, unknown>)).join('|')}`,
    )
  })
  return blocks
}

function settled(state: EditorState, accept: boolean): ProseMirrorNode {
  const tr = state.tr
  settleAllChanges(tr, accept)
  const doc = state.apply(tr).doc
  doc.check()
  return doc
}

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
    const next = state.apply(buildTrackedEdits(state, [{ range, markdown: 'cochila' }], [], MODEL, NOW)!)
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

describe('Markdown in a proposed edit', () => {
  beforeEach(() => useWorkspace.setState({ readOnly: false }))

  it('bold and italic become marks, inside the revision', () => {
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.'))))
    applyProposals(fake.editor, [edit(fake.state().doc, '1', 1, 'dorme', '**dorme** *bem*')], MODEL)
    expect(shown(fake.state().doc)).toEqual(['O gato [dorme]{dorme}{ }{bem}.'])
    const marks: string[][] = []
    fake.state().doc.descendants((node) => {
      if (node.isText)
        marks.push(node.marks.map((mark) => mark.type.name).filter((name) => name !== 'insertion'))
    })
    expect(marks).toEqual([[], ['deletion'], ['bold'], [], ['italic'], []])
  })

  it('keeps the spaces at the ends, which Markdown would drop', () => {
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.'))))
    applyProposals(fake.editor, [edit(fake.state().doc, '1', 1, 'gato ', 'cão ')], MODEL)
    expect(shown(fake.state().doc)).toEqual(['O [gato ]{cão }dorme.'])
  })

  it('a word fixed inside bold text keeps the bold', () => {
    const fake = fakeEditor(docOf(paragraph(text('Muito obrigdo.', { type: 'bold' }))))
    applyProposals(fake.editor, [edit(fake.state().doc, '1', 1, 'obrigdo', 'obrigado')], MODEL)
    const inserted = fake.state().doc.firstChild!.child(2)
    expect(inserted.text).toBe('obrigado')
    expect(inserted.marks.map((mark) => mark.type.name).sort()).toEqual(['bold', 'insertion'])
  })
})

describe('a whole paragraph that the answer turns into other blocks', () => {
  beforeEach(() => useWorkspace.setState({ readOnly: false }))

  it('becomes a heading; rejecting brings back the paragraph', () => {
    const fake = fakeEditor(docOf(paragraph(text('relatorio de vendas')), paragraph(text('Texto.'))))
    applyProposals(
      fake.editor,
      [edit(fake.state().doc, '1', 1, 'relatorio de vendas', '# Relatório de vendas')],
      MODEL,
    )
    expect(outline(settled(fake.state(), true))).toEqual(['heading1:Relatório de vendas', 'paragraph:Texto.'])
    expect(outline(settled(fake.state(), false))).toEqual([
      'paragraph:relatorio de vendas',
      'paragraph:Texto.',
    ])
  })

  it('becomes a list with nothing left over on accept', () => {
    const fake = fakeEditor(docOf(paragraph(text('a, b')), paragraph(text('Fim.'))))
    applyProposals(fake.editor, [edit(fake.state().doc, '1', 1, 'a, b', '1. a\n2. b')], MODEL)
    expect(outline(settled(fake.state(), true))).toEqual(['orderedList:a|b', 'paragraph:Fim.'])
    expect(outline(settled(fake.state(), false))).toEqual(['paragraph:a, b', 'paragraph:Fim.'])
  })

  it('plain text over a heading keeps the heading, which the model cannot see', () => {
    const heading = { type: 'heading', attrs: { level: 2 }, content: [text('Resumo')] }
    const fake = fakeEditor(docOf(heading))
    applyProposals(fake.editor, [edit(fake.state().doc, '1', 1, 'Resumo', 'Resumo final')], MODEL)
    expect(outline(settled(fake.state(), true))).toEqual(['heading2:Resumo final'])
  })

  it('part of a paragraph stays inside it', () => {
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.'))))
    applyProposals(fake.editor, [edit(fake.state().doc, '1', 1, 'gato', '# gato')], MODEL)
    expect(outline(settled(fake.state(), true))).toEqual(['paragraph:O gato dorme.'])
  })
})

describe('applying an insertion', () => {
  beforeEach(() => useWorkspace.setState({ readOnly: false }))

  const GUIDE = '# Guia\n\nTexto com **negrito**.\n\n- Café\n- Chá'

  it('writes headings, paragraphs and lists into an empty document, all as revision', () => {
    const fake = fakeEditor(docOf(paragraph()))
    const outcome = applyProposals(fake.editor, [insertion(fake.state().doc, '1', 0, GUIDE)], MODEL)
    expect(outcome).toEqual({ applied: ['1'], changed: [] })
    expect(outline(fake.state().doc)).toEqual([
      'heading1:{Guia}',
      'paragraph:{Texto com }{negrito}{.}',
      'bulletList:{Café}|{Chá}',
      'paragraph:',
    ])
    expect(authorsOf(fake.state().doc)).toEqual([MODEL])
  })

  it('rejecting it brings back the empty document; accepting it keeps the formatting', () => {
    const fake = fakeEditor(docOf(paragraph()))
    applyProposals(fake.editor, [insertion(fake.state().doc, '1', 1, GUIDE)], MODEL)
    expect(outline(settled(fake.state(), false))).toEqual(['paragraph:'])
    expect(outline(settled(fake.state(), true))).toEqual([
      'heading1:Guia',
      'paragraph:Texto com negrito.',
      'bulletList:Café|Chá',
      'paragraph:',
    ])
  })

  it('goes after its paragraph, and after the whole list when the paragraph is an item', () => {
    const fake = fakeEditor(docOf(paragraph(text('Um.')), bullets('a', 'b'), paragraph(text('Fim.'))))
    const doc = fake.state().doc
    applyProposals(
      fake.editor,
      [insertion(doc, '1', 1, 'Depois do um.'), insertion(doc, '2', 2, 'Depois da lista.')],
      MODEL,
    )
    expect(outline(fake.state().doc)).toEqual([
      'paragraph:Um.',
      'paragraph:{Depois do um.}',
      'bulletList:a|b',
      'paragraph:{Depois da lista.}',
      'paragraph:Fim.',
    ])
  })

  it('insertions at one place keep their order, in the same transaction as the edits', () => {
    const fake = fakeEditor(docOf(paragraph(text('O gato dorme.'))))
    const doc = fake.state().doc
    const outcome = applyProposals(
      fake.editor,
      [
        insertion(doc, '1', 1, 'Primeira.'),
        edit(doc, '2', 1, 'gato', 'cão'),
        insertion(doc, '3', 1, 'Segunda.'),
      ],
      MODEL,
    )
    expect(outcome).toEqual({ applied: ['1', '2', '3'], changed: [] })
    expect(fake.transactions).toHaveLength(1)
    expect(outline(fake.state().doc)).toEqual([
      'paragraph:O [gato]{cão} dorme.',
      'paragraph:{Primeira.}',
      'paragraph:{Segunda.}',
    ])
  })

  it('a paragraph that changed since the proposal: nothing goes in', () => {
    const proposal = insertion(docOf(paragraph(text('Antes.'))), '1', 1, 'Novo.')
    const fake = fakeEditor(docOf(paragraph(text('Depois.'))))
    expect(applyProposals(fake.editor, [proposal], MODEL)).toEqual({ applied: [], changed: ['1'] })
    expect(fake.transactions).toHaveLength(0)
  })
})
