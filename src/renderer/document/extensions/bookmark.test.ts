import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Fragment, Node as ProseMirrorNode, Slice } from '@tiptap/pm/model'
import { EditorState, TextSelection } from '@tiptap/pm/state'
import { buildEditorExtensions } from '../editor-extensions.js'
import {
  bookmarksOf,
  extendOverBookmarks,
  placeBookmark,
  removeBookmark,
  withoutRepeatedBookmarks,
} from './bookmark.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const start = (name: string, bid: string) => ({ type: 'bookmarkStart', attrs: { name, bid } })
const end = (bid: string) => ({ type: 'bookmarkEnd', attrs: { bid } })
const text = (value: string) => ({ type: 'text', text: value })

function docOf(...paragraphs: unknown[][]): ProseMirrorNode {
  return ProseMirrorNode.fromJSON(schema, {
    type: 'doc',
    content: paragraphs.map((content) => ({ type: 'paragraph', content })),
  })
}

describe('bookmarksOf', () => {
  it('casa as pontas pelo id, ainda que em parágrafos diferentes', () => {
    const doc = docOf([start('Resumo', '4'), text('um')], [text('dois'), end('4')])
    const [entry] = bookmarksOf(doc)
    expect(entry).toMatchObject({ name: 'Resumo', bid: '4', pos: 1 })
    expect(entry!.end).toBeGreaterThan(entry!.pos)
  })
})

describe('placeBookmark e removeBookmark', () => {
  it('marca a seleção com um id novo, e o nome repetido muda de lugar', () => {
    const doc = docOf([start('Velho', '2'), text('abcdef'), end('2')])
    let state = EditorState.create({ doc })
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 4)))

    const tr = state.tr
    placeBookmark(tr, 'Novo')
    const after = bookmarksOf(tr.doc)
    expect(after.map((bookmark) => [bookmark.name, bookmark.bid])).toEqual([
      ['Velho', '2'],
      ['Novo', '3'],
    ])

    // O mesmo nome outra vez: continua um só.
    placeBookmark(tr, 'Novo')
    expect(bookmarksOf(tr.doc).filter((bookmark) => bookmark.name === 'Novo')).toHaveLength(1)

    expect(removeBookmark(tr, 'Velho')).toBe(true)
    expect(bookmarksOf(tr.doc).map((bookmark) => bookmark.name)).toEqual(['Novo'])
    // O texto fica.
    expect(tr.doc.textContent).toBe('abcdef')
  })
})

describe('withoutRepeatedBookmarks', () => {
  it('tira da colagem o marcador que o documento já tem, e deixa o novo', () => {
    const doc = docOf([start('Resumo', '1'), text('x'), end('1')])
    const pasted = docOf([start('Resumo', '1'), text('cópia'), end('1')], [start('Outro', '9'), end('9')])
    const slice = new Slice(Fragment.from(pasted.content), 0, 0)

    const result = withoutRepeatedBookmarks(slice, doc)
    const names: string[] = []
    const kinds: string[] = []
    result.content.descendants((node) => {
      kinds.push(node.type.name)
      if (node.type.name === 'bookmarkStart') names.push(String(node.attrs['name']))
      return true
    })

    expect(names).toEqual(['Outro'])
    expect(kinds.filter((kind) => kind === 'bookmarkEnd')).toHaveLength(1)
  })

  it('renumera o marcador de nome novo cujo id o documento já usa', () => {
    const doc = docOf([start('Resumo', '0'), text('x'), end('0')])
    const pasted = docOf([start('Colado', '0'), text('y'), end('0')])
    const result = withoutRepeatedBookmarks(new Slice(Fragment.from(pasted.content), 0, 0), doc)
    const marks: string[] = []
    result.content.descendants((node) => {
      if (node.type.name.startsWith('bookmark')) marks.push(`${node.type.name}:${String(node.attrs['bid'])}`)
      return true
    })
    expect(marks).toEqual(['bookmarkStart:1', 'bookmarkEnd:1'])
  })

  it('arrastar dentro do documento move o marcador junto', () => {
    const doc = docOf([start('Resumo', '1'), text('x'), end('1')])
    const slice = new Slice(Fragment.from(doc.content), 0, 0)
    expect(withoutRepeatedBookmarks(slice, doc, true)).toBe(slice)
  })
})

describe('seleção de uma linha inteira', () => {
  it('leva junto os marcadores encostados no começo e no fim do texto', () => {
    const doc = docOf([start('_Ref1', '1'), text('Figura 1'), end('1')])
    // O começo do texto é a posição 2 (depois do marcador); o fim, 10.
    let state = EditorState.create({ doc })
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 10)))
    const extended = extendOverBookmarks(state)
    expect(extended).not.toBeNull()
    expect([extended!.selection.from, extended!.selection.to]).toEqual([1, 11])

    // No meio do texto, nada muda.
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3, 5)))
    expect(extendOverBookmarks(state)).toBeNull()
  })
})
