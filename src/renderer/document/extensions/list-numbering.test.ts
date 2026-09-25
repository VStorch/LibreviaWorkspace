import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode, type DOMOutputSpec } from '@tiptap/pm/model'
import { buildEditorExtensions } from '../editor-extensions.js'
import { drawListsForPrint } from './list-numbering.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const item = (text: string, ...nested: unknown[]) => ({
  type: 'listItem',
  content: [{ type: 'paragraph', content: [{ type: 'text', text }] }, ...nested],
})

/** Os atributos que o nó põe no elemento, sem precisar de DOM. */
function renderedAttrs(node: ProseMirrorNode): Record<string, unknown> {
  const spec = node.type.spec.toDOM!(node) as DOMOutputSpec
  const attrs = Array.isArray(spec) ? spec[1] : null
  return attrs !== null && typeof attrs === 'object' && !Array.isArray(attrs)
    ? (attrs as Record<string, unknown>)
    : {}
}

describe('drawListsForPrint', () => {
  it('leva ao papel a marca e o recuo que a tela desenha por decoração', () => {
    const doc = ProseMirrorNode.fromJSON(schema, {
      type: 'doc',
      content: [
        {
          type: 'orderedList',
          content: [item('um', { type: 'orderedList', content: [item('dentro')] }), item('dois')],
        },
      ],
    })

    const [list] = drawListsForPrint(doc)
    expect(renderedAttrs(list!)['data-list-indent']).toBe('')
    expect(String(renderedAttrs(list!)['style'])).toContain('--lista-recuo: 12.7mm')

    const first = list!.child(0)
    const nested = first.child(1)
    expect(renderedAttrs(first)['data-label']).toBe('1.')
    expect(renderedAttrs(nested.child(0))['data-label']).toBe('a.')
    expect(renderedAttrs(list!.child(1))['data-label']).toBe('2.')
  })

  it('o documento em si não ganha atributo nenhum', () => {
    const doc = ProseMirrorNode.fromJSON(schema, {
      type: 'doc',
      content: [{ type: 'bulletList', content: [item('um')] }],
    })
    drawListsForPrint(doc)
    expect(doc.child(0).attrs['listDraw']).toBeNull()
    expect(doc.child(0).child(0).attrs['listDraw']).toBeNull()
  })
})
