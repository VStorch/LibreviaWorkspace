import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { EditorState } from '@tiptap/pm/state'
import type { Decoration } from '@tiptap/pm/view'
import type { AiProposal } from '@services/ai/proposals.js'
import { buildEditorExtensions } from '../editor-extensions.js'
import { chatParagraphs } from '../ai-proposals.js'
import { aiDiffKey, aiDiffPlugin, type AiDiffEntry } from './ai-diff.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const paragraph = (value: string): Record<string, unknown> =>
  value === '' ? { type: 'paragraph' } : { type: 'paragraph', content: [{ type: 'text', text: value }] }

function stateOf(...paragraphs: string[]): EditorState {
  const doc = ProseMirrorNode.fromJSON(schema, { type: 'doc', content: paragraphs.map(paragraph) })
  return EditorState.create({ doc, plugins: [aiDiffPlugin()] })
}

const edit = (state: EditorState, find: string, replace: string): AiProposal => ({
  id: '1',
  kind: 'edit',
  paragraph: 1,
  paragraphText: chatParagraphs(state.doc)[0]!,
  find,
  replace,
})

function shown(state: EditorState, proposals: readonly AiProposal[]): EditorState {
  const entries: AiDiffEntry[] = proposals.map((proposal) => ({ key: `m:${proposal.id}`, proposal }))
  return state.apply(state.tr.setMeta(aiDiffKey, { entries, readOnly: false, decide: () => undefined }))
}

/** Each decoration as `from-to class` for inline ones, `at key` for widgets. */
function drawn(state: EditorState): string[] {
  const decorations = aiDiffKey.getState(state)!.decorations.find() as (Decoration & {
    type: { attrs?: { class?: string }; spec?: { key?: string } }
  })[]
  const byPlace = [...decorations].sort((a, b) => a.from - b.from || a.to - b.to)
  return byPlace.map((decoration) =>
    decoration.from === decoration.to
      ? `${decoration.from} ${String(decoration.spec['key']).split(':').slice(2, 4).join(':')}`
      : `${decoration.from}-${decoration.to} ${decoration.type.attrs?.class}`,
  )
}

describe('the assistant diff on the sheet', () => {
  it('a word changed inside a paragraph: the old word struck, the new one beside it, then the buttons', () => {
    const state = stateOf('O gato dorme.')
    const next = shown(state, [edit(state, 'gato dorme', 'gato cochila')])
    expect(drawn(next)).toEqual(['8-13 ai-diff__old', '13 2:cochila', '13 actions:false'])
  })

  it('an answer of several blocks goes as a block after the paragraph', () => {
    const state = stateOf('O gato dorme.', 'Fim.')
    const next = shown(state, [edit(state, 'O gato dorme.', '# Gato\n\n- dorme')])
    expect(drawn(next)).toEqual(['1-14 ai-diff__old', '15 block:# Gato\n\n- dorme'])
  })

  it('an insertion goes between the blocks, a comment marks its text', () => {
    const state = stateOf('O gato dorme.')
    const insert: AiProposal = {
      id: '2',
      kind: 'insert',
      paragraph: 1,
      paragraphText: 'O gato dorme.',
      content: 'Novo.',
    }
    const comment: AiProposal = {
      id: '3',
      kind: 'comment',
      paragraph: 1,
      paragraphText: 'O gato dorme.',
      find: 'gato',
      comment: 'Qual?',
    }
    const next = shown(state, [insert, comment])
    expect(drawn(next)).toEqual(['3-7 ai-diff__comment', '7 note:Qual?', '15 block:Novo.'])
  })

  it('follows the text as the person types, and a changed paragraph loses its diff', () => {
    const state = stateOf('O gato dorme.')
    const next = shown(state, [edit(state, 'dorme', 'cochila')])
    const typedBefore = next.apply(next.tr.insertText('X', 1))
    expect(drawn(typedBefore)).toEqual([])
    const typedAfter = next.apply(next.tr.insert(next.doc.content.size, schema.nodes['paragraph']!.create()))
    expect(drawn(typedAfter)).toEqual(['8-13 ai-diff__old', '13 1:cochila', '13 actions:false'])
  })

  it('no open proposal, nothing drawn', () => {
    const state = stateOf('O gato dorme.')
    expect(drawn(shown(shown(state, [edit(state, 'gato', 'cão')]), []))).toEqual([])
  })
})
