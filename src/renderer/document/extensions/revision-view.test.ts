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

// "ab" + [cd excluído] + {ef inserido} + "gh": posições 1..9 no primeiro parágrafo.
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

describe('o cursor fora do que o modo esconde', () => {
  it('cada modo esconde o seu trecho', () => {
    const block = doc.firstChild!
    expect(hiddenRuns(block, RevisionView.All)).toEqual([])
    expect(hiddenRuns(block, RevisionView.None)).toEqual([{ from: 2, to: 4 }])
    expect(hiddenRuns(block, RevisionView.Original)).toEqual([{ from: 4, to: 6 }])
  })

  it('dentro do excluído escondido, o cursor vai para a borda no sentido em que andava', () => {
    expect(visiblePosition(doc, 4, RevisionView.Simple, 1)).toBe(5)
    expect(visiblePosition(doc, 4, RevisionView.Simple, -1)).toBe(3)
    // Nas bordas já está fora; e na marcação completa nada se esconde.
    expect(visiblePosition(doc, 3, RevisionView.Simple, 1)).toBe(3)
    expect(visiblePosition(doc, 4, RevisionView.All, 1)).toBe(4)
  })

  it('o parágrafo excluído inteiro some, e o cursor passa para o vizinho', () => {
    const second = doc.child(1)
    expect(isHiddenBlock(second, RevisionView.None)).toBe(true)
    expect(isHiddenBlock(second, RevisionView.Original)).toBe(false)
    const inside = doc.firstChild!.nodeSize + 3
    const after = visiblePosition(doc, inside, RevisionView.None, 1)
    expect(doc.resolve(after).parent.textContent).toBe('fim')
    const before = visiblePosition(doc, inside, RevisionView.None, -1)
    expect(doc.resolve(before).parent.textContent).toBe('abcdefgh')
  })

  it('a seleção que já está fora fica como está', () => {
    expect(visibleSelection(doc, TextSelection.create(doc, 2), RevisionView.None, 1)).toBeNull()
    const moved = visibleSelection(doc, TextSelection.create(doc, 6), RevisionView.Original, -1)
    expect(moved?.head).toBe(5)
  })
})

describe('apagar o que se vê', () => {
  it('a seleção que passa pelo excluído escondido apaga só o resto', () => {
    const tr = EditorState.create({ doc }).tr
    // "ab[cd]ef gh" de 1 a 9: sem marcação, "cd" não se vê e fica.
    deleteVisible(tr, 1, 9, RevisionView.None)
    expect(tr.doc.child(0).textContent).toBe('cd')
    // Com tudo à vista, sai tudo.
    const all = deleteVisible(EditorState.create({ doc }).tr, 1, 9, RevisionView.All)
    expect(all.doc.child(0).textContent).toBe('')
  })
})

describe('o cursor fora de muitos parágrafos escondidos', () => {
  it('passa por mais de dezesseis em sequência', () => {
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
