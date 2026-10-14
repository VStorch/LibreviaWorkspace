import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { AllSelection, EditorState, NodeSelection, TextSelection } from '@tiptap/pm/state'
import { buildEditorExtensions } from './editor-extensions.js'
import { aiTarget, buildInsertBelow, buildTrackedReplacement, rangeChanged } from './ai-apply.js'
import { SKIP_TRACKING } from './extensions/track-input.js'

const schema = getSchema(buildEditorExtensions(() => {}))
const NOW = new Date('2026-10-01T12:34:56Z')

const paragraph = (...content: Record<string, unknown>[]): Record<string, unknown> => ({
  type: 'paragraph',
  content,
})
const text = (value: string, ...marks: Record<string, unknown>[]): Record<string, unknown> =>
  marks.length === 0 ? { type: 'text', text: value } : { type: 'text', text: value, marks }

function stateOf(blocks: Record<string, unknown>[], from: number, to: number): EditorState {
  const doc = ProseMirrorNode.fromJSON(schema, { type: 'doc', content: blocks })
  const state = EditorState.create({ doc })
  return state.apply(state.tr.setSelection(TextSelection.create(doc, from, to)))
}

function shown(doc: ProseMirrorNode): string[] {
  const blocks: string[] = []
  doc.descendants((node) => {
    if (!node.isTextblock) return true
    let line = ''
    node.forEach((child) => {
      const value = child.text ?? ''
      if (child.marks.some((m) => m.type.name === 'deletion')) line += `[${value}]`
      else if (child.marks.some((m) => m.type.name === 'insertion')) line += `{${value}}`
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
      if (mark.type.name === 'insertion' || mark.type.name === 'deletion')
        authors.add(String(mark.attrs['author']))
    }
  })
  return [...authors]
}

describe('aiTarget', () => {
  it('devolve o trecho selecionado como texto puro', () => {
    const state = stateOf([paragraph(text('O gato dorme.'))], 3, 7)
    expect(aiTarget(state)).toEqual({ range: { from: 3, to: 7 }, text: 'gato' })
  })

  it('deixa de fora o texto já excluído numa revisão', () => {
    const deletion = { type: 'deletion', attrs: { author: 'Ana', date: '2026-01-01T00:00:00Z', rid: '1' } }
    const state = stateOf([paragraph(text('um '), text('dois ', deletion), text('três'))], 1, 13)
    expect(aiTarget(state)).toMatchObject({ text: 'um três' })
  })

  it('recusa seleção vazia, com imagem e entre células', () => {
    expect(aiTarget(stateOf([paragraph(text('abc'))], 2, 2))).toEqual({ problem: 'empty' })

    const image = { type: 'image', attrs: { src: 'data:image/png;base64,AA==' } }
    expect(aiTarget(stateOf([paragraph(text('a'), image, text('b'))], 1, 4))).toEqual({
      problem: 'unsupported',
    })

    const cell = (value: string): Record<string, unknown> => ({
      type: 'tableCell',
      content: [paragraph(text(value))],
    })
    const table = { type: 'table', content: [{ type: 'tableRow', content: [cell('um'), cell('dois')] }] }
    const doc = ProseMirrorNode.fromJSON(schema, { type: 'doc', content: [table] })
    const across = EditorState.create({ doc }).apply(
      EditorState.create({ doc }).tr.setSelection(TextSelection.create(doc, 4, 11)),
    )
    expect(aiTarget(across)).toEqual({ problem: 'unsupported' })
  })

  it('selecionar tudo vira o texto do primeiro ao último parágrafo', () => {
    const doc = ProseMirrorNode.fromJSON(schema, {
      type: 'doc',
      content: [paragraph(text('Um')), paragraph(text('Dois'))],
    })
    const state = EditorState.create({ doc })
    const all = state.apply(state.tr.setSelection(new AllSelection(doc)))
    expect(aiTarget(all)).toEqual({ range: { from: 1, to: 9 }, text: 'Um\nDois' })
    const next = all.apply(buildTrackedReplacement(all, { from: 1, to: 9 }, 'Três', 'm', NOW)!)
    expect(shown(next.doc)).toEqual(['[Um]', '[Dois]{Três}'])
  })

  it('recusa seleção de nó inteiro', () => {
    const image = { type: 'image', attrs: { src: 'data:image/png;base64,AA==' } }
    const doc = ProseMirrorNode.fromJSON(schema, { type: 'doc', content: [paragraph(text('a'), image)] })
    const state = EditorState.create({ doc })
    expect(aiTarget(state.apply(state.tr.setSelection(NodeSelection.create(doc, 2))))).toEqual({
      problem: 'unsupported',
    })
  })
})

describe('buildTrackedReplacement', () => {
  it('entra como revisão com o modelo de autor, mesmo com o controle desligado', () => {
    const state = stateOf([paragraph(text('O gato dorme.'))], 3, 7)
    const tr = buildTrackedReplacement(state, { from: 3, to: 7 }, 'cão', 'llama3', NOW)!
    expect(tr.getMeta(SKIP_TRACKING)).toBe(true)
    const next = state.apply(tr)
    expect(shown(next.doc)).toEqual(['O [gato]{cão} dorme.'])
    expect(authorsOf(next.doc)).toEqual(['llama3'])
  })

  it('várias linhas viram parágrafos ligados ao texto em volta', () => {
    const state = stateOf([paragraph(text('Antes meio depois'))], 7, 11)
    const next = state.apply(buildTrackedReplacement(state, { from: 7, to: 11 }, 'um\ndois', 'm', NOW)!)
    next.doc.check()
    expect(shown(next.doc)).toEqual(['Antes [meio]{um}', '{dois} depois'])
  })

  it('mantém a formatação do começo do trecho', () => {
    const state = stateOf([paragraph(text('negrito', { type: 'bold' }))], 1, 8)
    const next = state.apply(buildTrackedReplacement(state, { from: 1, to: 8 }, 'forte', 'm', NOW)!)
    let bold = false
    next.doc.descendants((node) => {
      if (node.text === 'forte') bold = node.marks.some((mark) => mark.type.name === 'bold')
    })
    expect(bold).toBe(true)
  })

  it('não faz nada com resposta vazia', () => {
    const state = stateOf([paragraph(text('abc'))], 1, 3)
    expect(buildTrackedReplacement(state, { from: 1, to: 3 }, '', 'm', NOW)).toBeNull()
  })
})

describe('buildInsertBelow', () => {
  it('insere parágrafos marcados depois do parágrafo do trecho', () => {
    const state = stateOf([paragraph(text('Primeiro')), paragraph(text('Segundo'))], 1, 5)
    const next = state.apply(buildInsertBelow(state, { from: 1, to: 5 }, 'Resumo', 'm', NOW)!)
    expect(shown(next.doc)).toEqual(['Primeiro', '{Resumo}', 'Segundo'])
  })
})

describe('rangeChanged', () => {
  it('acusa trecho editado ou colapsado', () => {
    const state = stateOf([paragraph(text('O gato dorme.'))], 3, 7)
    expect(rangeChanged(state.doc, { from: 3, to: 7 }, 'gato')).toBe(false)
    expect(rangeChanged(state.doc, { from: 3, to: 7 }, 'rato')).toBe(true)
    expect(rangeChanged(state.doc, { from: 5, to: 5 }, '')).toBe(true)
  })
})
