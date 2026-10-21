import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { AllSelection, EditorState, NodeSelection, TextSelection } from '@tiptap/pm/state'
import { buildEditorExtensions } from './editor-extensions.js'
import { AiFormat } from '@shared/types.js'
import {
  aiTarget,
  buildInsertBelow,
  buildTrackedReplacement,
  chatDocument,
  documentAround,
  rangeChanged,
  selectedMarkdown,
} from './ai-apply.js'
import { SKIP_TRACKING } from './extensions/track-input.js'
import { settleAllChanges } from './extensions/track-changes.js'

const schema = getSchema(buildEditorExtensions(() => {}))
const NOW = new Date('2026-10-01T12:34:56Z')
const plain = (value: string) => ({ text: value, format: AiFormat.Plain })
const markdown = (value: string) => ({ text: value, format: AiFormat.Markdown })

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
  it('returns the selected excerpt as plain text', () => {
    const state = stateOf([paragraph(text('O gato dorme.'))], 3, 7)
    expect(aiTarget(state)).toEqual({ range: { from: 3, to: 7 }, text: 'gato' })
  })

  it('leaves out text already deleted in a revision', () => {
    const deletion = { type: 'deletion', attrs: { author: 'Ana', date: '2026-01-01T00:00:00Z', rid: '1' } }
    const state = stateOf([paragraph(text('um '), text('dois ', deletion), text('três'))], 1, 13)
    expect(aiTarget(state)).toMatchObject({ text: 'um três' })
  })

  it('rejects an empty selection, one with an image, and one across cells', () => {
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

  it('select all becomes the text from the first to the last paragraph', () => {
    const doc = ProseMirrorNode.fromJSON(schema, {
      type: 'doc',
      content: [paragraph(text('Um')), paragraph(text('Dois'))],
    })
    const state = EditorState.create({ doc })
    const all = state.apply(state.tr.setSelection(new AllSelection(doc)))
    expect(aiTarget(all)).toEqual({ range: { from: 1, to: 9 }, text: 'Um\nDois' })
    const next = all.apply(buildTrackedReplacement(all, { from: 1, to: 9 }, plain('Três'), 'm', NOW)!)
    expect(shown(next.doc)).toEqual(['[Um]', '[Dois]{Três}'])
  })

  it('rejects a whole-node selection', () => {
    const image = { type: 'image', attrs: { src: 'data:image/png;base64,AA==' } }
    const doc = ProseMirrorNode.fromJSON(schema, { type: 'doc', content: [paragraph(text('a'), image)] })
    const state = EditorState.create({ doc })
    expect(aiTarget(state.apply(state.tr.setSelection(NodeSelection.create(doc, 2))))).toEqual({
      problem: 'unsupported',
    })
  })
})

describe('buildTrackedReplacement', () => {
  it('goes in as a revision with the model as author, even with tracking off', () => {
    const state = stateOf([paragraph(text('O gato dorme.'))], 3, 7)
    const tr = buildTrackedReplacement(state, { from: 3, to: 7 }, plain('cão'), 'llama3', NOW)!
    expect(tr.getMeta(SKIP_TRACKING)).toBe(true)
    const next = state.apply(tr)
    expect(shown(next.doc)).toEqual(['O [gato]{cão} dorme.'])
    expect(authorsOf(next.doc)).toEqual(['llama3'])
  })

  it('several lines become paragraphs joined to the surrounding text', () => {
    const state = stateOf([paragraph(text('Antes meio depois'))], 7, 11)
    const next = state.apply(
      buildTrackedReplacement(state, { from: 7, to: 11 }, plain('um\ndois'), 'm', NOW)!,
    )
    next.doc.check()
    expect(shown(next.doc)).toEqual(['Antes [meio]{um}', '{dois} depois'])
  })

  it('keeps the formatting from the start of the excerpt', () => {
    const state = stateOf([paragraph(text('negrito', { type: 'bold' }))], 1, 8)
    const next = state.apply(buildTrackedReplacement(state, { from: 1, to: 8 }, plain('forte'), 'm', NOW)!)
    let bold = false
    next.doc.descendants((node) => {
      if (node.text === 'forte') bold = node.marks.some((mark) => mark.type.name === 'bold')
    })
    expect(bold).toBe(true)
  })

  it('does nothing with an empty answer', () => {
    const state = stateOf([paragraph(text('abc'))], 1, 3)
    expect(buildTrackedReplacement(state, { from: 1, to: 3 }, plain(''), 'm', NOW)).toBeNull()
  })
})

describe('Markdown answer', () => {
  it("bold and list become formatting inside the model's revision", () => {
    const state = stateOf(
      [paragraph(text('Antes')), paragraph(text('troque')), paragraph(text('Depois'))],
      8,
      14,
    )
    const tr = buildTrackedReplacement(state, { from: 8, to: 14 }, markdown('**x**\n\n- item'), 'm', NOW)!
    const next = state.apply(tr)
    next.doc.check()
    let strong = false
    next.doc.descendants((node) => {
      if (node.text === 'x') {
        strong = node.marks.some((mark) => mark.type.name === 'bold')
        expect(node.marks.some((mark) => mark.type.name === 'insertion')).toBe(true)
      }
    })
    expect(strong).toBe(true)
    expect(next.doc.child(2).type.name).toBe('bulletList')
    expect(shown(next.doc)).toContain('{item}')
    expect(authorsOf(next.doc)).toEqual(['m'])
  })

  it("the answer does not inherit the excerpt's bold, but inherits its color", () => {
    const red = { type: 'textStyle', attrs: { color: '#ff0000' } }
    const state = stateOf([paragraph(text('forte', { type: 'bold' }, red))], 1, 6)
    const next = state.apply(
      buildTrackedReplacement(state, { from: 1, to: 6 }, markdown('*leve*'), 'm', NOW)!,
    )
    next.doc.descendants((node) => {
      if (node.text !== 'leve') return
      const names = node.marks.map((mark) => mark.type.name)
      expect(names).toContain('italic')
      expect(names).toContain('textStyle')
      expect(names).not.toContain('bold')
    })
  })

  it('inserting below puts heading and list marked as insertion', () => {
    const state = stateOf([paragraph(text('Primeiro'))], 1, 5)
    const next = state.apply(
      buildInsertBelow(state, { from: 1, to: 5 }, markdown('# Resumo\n\n1. um'), 'm', NOW)!,
    )
    next.doc.check()
    expect(next.doc.child(1).type.name).toBe('heading')
    expect(next.doc.child(2).type.name).toBe('orderedList')
    expect(shown(next.doc)).toEqual(['Primeiro', '{Resumo}', '{um}'])
  })

  it('the excerpt goes as Markdown and the surrounding document without deletions', () => {
    const deleted = { type: 'deletion', attrs: { author: 'a', date: NOW.toISOString(), id: 1 } }
    const state = stateOf(
      [paragraph(text('Antes '), text('apagado', deleted), text('forte', { type: 'bold' }), text(' depois'))],
      14,
      19,
    )
    expect(selectedMarkdown(state.doc, { from: 14, to: 19 })).toBe('**forte**')
    expect(documentAround(state.doc, { from: 14, to: 19 })).toEqual({ before: 'Antes ', after: ' depois' })
  })
})

describe('Markdown answer under tracking', () => {
  const THREE = [paragraph(text('Primeiro.')), paragraph(text('O gato dorme.')), paragraph(text('Fim.'))]
  const ANSWER = markdown('O **gato** dorme.\n\n- um\n- dois')

  function settled(state: EditorState, accept: boolean): ProseMirrorNode {
    const tr = state.tr
    settleAllChanges(tr, accept)
    const doc = state.apply(tr).doc
    doc.check()
    return doc
  }

  function outline(doc: ProseMirrorNode): string[] {
    const blocks: string[] = []
    doc.forEach((node) => blocks.push(`${node.type.name}:${node.textContent}`))
    return blocks
  }

  it('code and code block go in as plain text, all marked as insertion', () => {
    const state = stateOf([paragraph(text('troque'))], 1, 7)
    const answer = markdown('usa `x`\n\n```\ny = 1\n```')
    const next = state.apply(buildTrackedReplacement(state, { from: 1, to: 7 }, answer, 'm', NOW)!)
    next.doc.check()
    next.doc.descendants((node) => {
      if (!node.isText) return
      const names = node.marks.map((mark) => mark.type.name)
      if (!names.includes('deletion')) expect(names).toContain('insertion')
      if (!names.includes('deletion')) expect(names).not.toContain('code')
    })
    expect(shown(next.doc)).toEqual(['[troque]{usa x}', '{y = 1}'])
  })

  it('rejecting paragraph and list returns only the original paragraphs', () => {
    const state = stateOf(THREE, 12, 25)
    const next = state.apply(buildTrackedReplacement(state, { from: 12, to: 25 }, ANSWER, 'm', NOW)!)
    expect(outline(settled(next, false))).toEqual([
      'paragraph:Primeiro.',
      'paragraph:O gato dorme.',
      'paragraph:Fim.',
    ])
  })

  it('accepting paragraph and list at the end of the excerpt leaves no empty paragraph', () => {
    const state = stateOf(THREE, 12, 25)
    const next = state.apply(buildTrackedReplacement(state, { from: 12, to: 25 }, ANSWER, 'm', NOW)!)
    expect(outline(settled(next, true))).toEqual([
      'paragraph:Primeiro.',
      'paragraph:O gato dorme.',
      'bulletList:umdois',
      'paragraph:Fim.',
    ])
  })

  it('list in the middle of a paragraph: accepting splits the paragraph and rejecting restores it', () => {
    const state = stateOf([paragraph(text('Antes troque depois.'))], 7, 13)
    const next = state.apply(buildTrackedReplacement(state, { from: 7, to: 13 }, markdown('- um'), 'm', NOW)!)
    expect(outline(settled(next, true))).toEqual(['paragraph:Antes ', 'bulletList:um', 'paragraph: depois.'])
    expect(outline(settled(next, false))).toEqual(['paragraph:Antes troque depois.'])
  })

  it('list at the start of the excerpt goes before the paragraph and vanishes on reject', () => {
    const state = stateOf(THREE, 12, 25)
    const answer = markdown('- um\n\nO gato dorme.')
    const next = state.apply(buildTrackedReplacement(state, { from: 12, to: 25 }, answer, 'm', NOW)!)
    expect(outline(settled(next, true))).toEqual([
      'paragraph:Primeiro.',
      'bulletList:um',
      'paragraph:O gato dorme.',
      'paragraph:Fim.',
    ])
    expect(outline(settled(next, false))).toEqual(outline(state.doc))
  })

  it('the Markdown excerpt carries superscript as text and, in a cell, only the paragraphs', () => {
    const state = stateOf([paragraph(text('m'), text('2', { type: 'superscript' }))], 1, 3)
    expect(selectedMarkdown(state.doc, { from: 1, to: 3 })).toBe('m2')
    const cell = { type: 'tableCell', content: [paragraph(text('um')), paragraph(text('dois'))] }
    const table = { type: 'table', content: [{ type: 'tableRow', content: [cell] }] }
    const inTable = stateOf([table], 4, 13)
    expect(selectedMarkdown(inTable.doc, { from: 4, to: 13 })).toBe('um\n\ndois')
  })
})

describe('buildInsertBelow', () => {
  it("inserts marked paragraphs after the excerpt's paragraph", () => {
    const state = stateOf([paragraph(text('Primeiro')), paragraph(text('Segundo'))], 1, 5)
    const next = state.apply(buildInsertBelow(state, { from: 1, to: 5 }, plain('Resumo'), 'm', NOW)!)
    expect(shown(next.doc)).toEqual(['Primeiro', '{Resumo}', 'Segundo'])
  })
})

describe('rangeChanged', () => {
  it('flags an edited or collapsed excerpt', () => {
    const state = stateOf([paragraph(text('O gato dorme.'))], 3, 7)
    expect(rangeChanged(state.doc, { from: 3, to: 7 }, 'gato')).toBe(false)
    expect(rangeChanged(state.doc, { from: 3, to: 7 }, 'rato')).toBe(true)
    expect(rangeChanged(state.doc, { from: 5, to: 5 }, '')).toBe(true)
  })
})

describe('chatDocument', () => {
  it('with no selection sends the whole document, without marking an excerpt', () => {
    const state = stateOf([paragraph(text('Um.')), paragraph(text('Dois.'))], 2, 2)
    expect(chatDocument(state)).toEqual({ before: 'Um.\nDois.', after: '', selection: '' })
  })

  it('with a selection sends the surrounding text and the selected excerpt', () => {
    const state = stateOf([paragraph(text('Antes meio depois'))], 7, 11)
    expect(chatDocument(state)).toEqual({ before: 'Antes ', after: ' depois', selection: 'meio' })
  })
})
