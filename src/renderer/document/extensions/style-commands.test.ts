import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { EditorState, TextSelection } from '@tiptap/pm/state'
import { BUILTIN_STYLES, type StyleSheet } from '@services/document/styles.js'
import { buildEditorExtensions } from '../editor-extensions.js'
import {
  applyCharacterStyle,
  applyParagraphStyle,
  clearDirectFormatting,
  markVisiblyOn,
  splitWithNextStyle,
  toggleInheritedMark,
} from './style-commands.js'

const schema = getSchema(buildEditorExtensions(() => {}))

/** Um título negrito pelo estilo: é onde o "desligado" tem de existir. */
const sheet: StyleSheet = {
  ...BUILTIN_STYLES,
  styles: {
    ...BUILTIN_STYLES.styles,
    Heading1: { ...BUILTIN_STYLES.styles['Heading1']!, character: { bold: true, fontSize: '16pt' } },
  },
}

function stateOf(doc: unknown, from = 1, to = from): EditorState {
  const state = EditorState.create({ schema, doc: schema.nodeFromJSON(doc) })
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)))
}

const paragraph = (text: string, attrs: Record<string, unknown> = {}, marks: unknown[] = []) => ({
  type: 'paragraph',
  attrs,
  content: [{ type: 'text', text, ...(marks.length === 0 ? {} : { marks }) }],
})

function run(state: EditorState, command: (tr: EditorState['tr']) => boolean): EditorState {
  const tr = state.tr
  expect(command(tr)).toBe(true)
  return state.apply(tr)
}

describe('comandos de estilo', () => {
  it('aplicar um estilo de título troca o bloco e tira a formatação direta', () => {
    const state = stateOf({ type: 'doc', content: [paragraph('Relatório', { spaceAfter: 20 })] })
    const after = run(state, (tr) => applyParagraphStyle(tr, sheet, 'Heading1'))
    const block = after.doc.firstChild!

    expect(block.type.name).toBe('heading')
    expect(block.attrs).toMatchObject({ level: 1, styleId: 'Heading1', spaceAfter: null })

    // E o padrão volta a parágrafo sem id — o que o leitor produz para o Normal.
    const back = run(after, (tr) => applyParagraphStyle(tr, sheet, 'Normal')).doc.firstChild!
    expect(back.type.name).toBe('paragraph')
    expect(back.attrs['styleId']).toBeNull()
  })

  it('limpar a formatação deixa o estilo e o link', () => {
    const link = { type: 'link', attrs: { href: 'https://exemplo.org' } }
    const state = stateOf(
      {
        type: 'doc',
        content: [
          paragraph('Texto', { styleId: 'ListParagraph', textAlign: 'center', indent: 2 }, [
            { type: 'bold' },
            link,
          ]),
        ],
      },
      2,
    )
    const block = run(state, clearDirectFormatting).doc.firstChild!

    expect(block.attrs).toMatchObject({ styleId: 'ListParagraph', textAlign: null, indent: 0 })
    expect(block.firstChild!.marks.map((mark) => mark.type.name)).toEqual(['link'])
  })

  it('tirar o negrito num título negrito grava a marca "desligado", e pô-lo de volta a tira', () => {
    const state = stateOf(
      {
        type: 'doc',
        content: [{ type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Título' }] }],
      },
      1,
      7,
    )
    expect(markVisiblyOn(state.tr, sheet, 'bold')).toBe(true)

    const off = run(state, (tr) => toggleInheritedMark(tr, sheet, 'bold'))
    expect(off.doc.firstChild!.firstChild!.marks[0]?.attrs['off']).toBe(true)
    expect(markVisiblyOn(off.tr, sheet, 'bold')).toBe(false)

    const on = run(off, (tr) => toggleInheritedMark(tr, sheet, 'bold'))
    expect(on.doc.firstChild!.firstChild!.marks).toEqual([])

    // Onde o estilo não liga, a alternância é a de sempre: quem a faz é o Tiptap.
    const plain = stateOf({ type: 'doc', content: [paragraph('Corpo')] }, 1, 6)
    expect(toggleInheritedMark(plain.tr, sheet, 'bold')).toBe(false)
  })

  it('Enter no fim do título abre um parágrafo Normal', () => {
    const state = stateOf(
      {
        type: 'doc',
        content: [
          {
            type: 'heading',
            attrs: { level: 1, styleId: 'Heading1' },
            content: [{ type: 'text', text: 'Fim' }],
          },
        ],
      },
      4,
    )
    const after = run(state, (tr) => splitWithNextStyle(tr, sheet))

    expect(after.doc.childCount).toBe(2)
    expect(after.doc.child(1).type.name).toBe('paragraph')
    expect(after.doc.child(1).attrs['styleId']).toBeNull()

    // No meio do título, o Enter é o de sempre.
    expect(splitWithNextStyle(stateOf(state.doc.toJSON(), 2).tr, sheet)).toBe(false)
  })

  it('o estilo de caractere vira marca no trecho', () => {
    const state = stateOf({ type: 'doc', content: [paragraph('Palavra')] }, 1, 8)
    const after = run(state, (tr) => applyCharacterStyle(tr, 'Hyperlink'))
    expect(after.doc.firstChild!.firstChild!.marks[0]?.attrs).toEqual({ styleId: 'Hyperlink' })
  })
})
