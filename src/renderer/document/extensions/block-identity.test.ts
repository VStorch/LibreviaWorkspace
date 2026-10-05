import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { splitBlock } from '@tiptap/pm/commands'
import { EditorState, TextSelection, type Transaction } from '@tiptap/pm/state'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { buildEditorExtensions } from '../editor-extensions.js'
import { uniqueOids } from './block-identity.js'

/**
 * Block identity is unique: with two blocks sharing an `oid`, the writer keeps the first one's XML
 * and regenerates the second. Enter splits the paragraph, and the new side does not take the `oid`.
 * Without Tiptap's `Editor`, which needs a DOM: Enter is ProseMirror's `splitBlock`.
 */

const schema = getSchema(buildEditorExtensions(() => {}))

function stateOf(doc: unknown): EditorState {
  return EditorState.create({ doc: schema.nodeFromJSON(doc), plugins: [uniqueOids()] })
}

/** Dispatch goes through `apply`, as in the editor. */
function run(state: EditorState, at: number, command: typeof splitBlock): EditorState {
  let next = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at)))
  command(next, (tr: Transaction) => {
    next = next.apply(tr)
  })
  return next
}

const oidsOf = (doc: ProseMirrorNode): unknown[] =>
  doc.children.map((block: ProseMirrorNode) => block.attrs['oid'])

describe('uniqueOids', () => {
  it('dividir um parágrafo deixa o `oid` com o lado de cima', () => {
    const state = stateOf({
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          attrs: { oid: 'b1' },
          content: [{ type: 'text', text: 'Antes e depois' }],
        },
      ],
    })

    // The cursor between "Antes" and " e depois": 1 is the start of the paragraph.
    const divided = run(state, 1 + 'Antes'.length, splitBlock)

    expect(divided.doc.childCount).toBe(2)
    expect(oidsOf(divided.doc)).toEqual(['b1', null])
  })

  it('a identidade repetida por colagem também sai', () => {
    // Saving already handled a repeated `oid` by keeping the first occurrence's XML and
    // regenerating the others. The rule is the same here, one step earlier: the second occurrence
    // stops claiming an identity that is not its own.
    const state = stateOf({
      type: 'doc',
      content: [
        { type: 'paragraph', attrs: { oid: 'b1' }, content: [{ type: 'text', text: 'Original' }] },
        { type: 'paragraph', attrs: { oid: 'b1' }, content: [{ type: 'text', text: 'Colado' }] },
        { type: 'paragraph', attrs: { oid: 'b2' }, content: [{ type: 'text', text: 'Outro' }] },
      ],
    })

    const typed = state.apply(state.tr.insertText('!', 1))

    expect(oidsOf(typed.doc)).toEqual(['b1', null, 'b2'])
  })

  it('não mexe no que já é único', () => {
    // The plugin returns a transaction only when there is something to fix: an appended transaction
    // on every key press dirties the undo history.
    const state = stateOf({
      type: 'doc',
      content: [
        { type: 'paragraph', attrs: { oid: 'b1' }, content: [{ type: 'text', text: 'Um' }] },
        { type: 'paragraph', attrs: { oid: 'b2' }, content: [{ type: 'text', text: 'Dois' }] },
      ],
    })

    const typed = state.apply(state.tr.insertText('!', 1))

    expect(oidsOf(typed.doc)).toEqual(['b1', 'b2'])
  })
})
