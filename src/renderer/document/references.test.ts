import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { BUILTIN_STYLES } from '@services/document/styles.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { EditorState, TextSelection, type Transaction } from '@tiptap/pm/state'
import type { Editor } from '@tiptap/react'
import { DEFAULT_PAGE_SETUP } from '@services/document/model.js'
import {
  captionLabels,
  crossReferenceTargets,
  insertCrossReference,
  noteNumberIn,
  settlePageFields,
  sheetAt,
  updateFields,
  type ReferenceContext,
} from './references.js'
import type { PageLayout } from './usePagination.js'

const schema = getSchema(buildEditorExtensions(() => {}))

const paragraph = (...content: unknown[]) => ({ type: 'paragraph', content })
const text = (value: string) => ({ type: 'text', text: value })
const seq = (label: string, result: string) => ({
  type: 'field',
  attrs: { instr: ` SEQ ${label} \\* ARABIC `, result },
})

const doc = ProseMirrorNode.fromJSON(schema, {
  type: 'doc',
  content: [
    { type: 'heading', attrs: { level: 1 }, content: [text('Um')] },
    paragraph(text('Figura '), seq('Figura', '1'), text(' — A')),
    paragraph(text('Tabela '), seq('Tabela', '1')),
    { type: 'heading', attrs: { level: 2 }, content: [text('Dois')] },
  ],
})

describe('sheetAt', () => {
  it('conta as folhas que começam antes da posição', () => {
    // The second sheet opens at the third block; the third, in the middle of the fourth.
    const starts = [{ blockIndex: 2 }, { blockIndex: 3, offset: 2 }]
    let third = 0
    let fourth = 0
    doc.forEach((_node, pos, index) => {
      if (index === 2) third = pos
      if (index === 3) fourth = pos
    })
    expect(sheetAt(doc, starts, 0)).toBe(1)
    expect(sheetAt(doc, starts, third)).toBe(2)
    expect(sheetAt(doc, starts, fourth)).toBe(2)
    expect(sheetAt(doc, starts, fourth + 3)).toBe(3)
  })
})

describe('destinos das referências', () => {
  it('acha as legendas pelo rótulo do SEQ e os títulos com o recuo do nível', () => {
    expect(captionLabels(doc, ['Figura', 'Equação'])).toEqual(['Figura', 'Equação', 'Tabela'])
    expect(
      crossReferenceTargets(doc, BUILTIN_STYLES, { type: 'caption', label: 'figura' }).map((t) => t.text),
    ).toEqual(['Figura 1 — A'])
    expect(crossReferenceTargets(doc, BUILTIN_STYLES, { type: 'heading' }).map((t) => t.text)).toEqual([
      'Um',
      ' Dois',
    ])
  })
})

/** A fake editor: the state and `dispatch`, which is all these functions use. */
function fakeEditor(json: unknown): { editor: Editor; fields: () => string[] } {
  let state = EditorState.create({ doc: ProseMirrorNode.fromJSON(schema, json) })
  const editor = {
    get state() {
      return state
    },
    view: { dispatch: (tr: Transaction) => (state = state.apply(tr)) },
  } as unknown as Editor
  const fields = (): string[] => {
    const results: string[] = []
    state.doc.descendants((node) => {
      if (node.type.name === 'field') results.push(String(node.attrs['result']))
      return true
    })
    return results
  }
  return { editor, fields }
}

function contextWith(layout: Partial<PageLayout>, outsideBookmarks: string[] = []): ReferenceContext {
  return {
    layout: {
      pages: 2,
      stackHeightPx: 0,
      sheetTops: [],
      sheetHeights: [],
      pageStarts: [],
      anchors: [],
      sheets: [],
      sheetWidths: [],
      stackWidthPx: 0,
      contentSheets: [],
      columnMoves: [],
      columnLines: [],
      noteAreas: [],
      ...layout,
    },
    page: DEFAULT_PAGE_SETUP,
    styles: BUILTIN_STYLES,
    setStyles: () => {},
    t: (key) => key,
    outsideBookmarks,
  }
}

const field = (instr: string, result: string) => ({ type: 'field', attrs: { instr, result } })

describe('atualizar campos', () => {
  it('a referência a marcador que o arquivo tem fora dos nós fica como estava', () => {
    const { editor, fields } = fakeEditor({
      type: 'doc',
      content: [paragraph(field(' REF Linhas \\h ', 'texto do Word'), field(' REF Sumiu ', 'velho'))],
    })
    updateFields(editor, contextWith({}, ['Linhas']))
    expect(fields()).toEqual(['texto do Word', 'references.field.missingBookmark'])
  })

  it('F9 que não mudou nada não arma o segundo passe', () => {
    const { editor, fields } = fakeEditor({
      type: 'doc',
      content: [
        paragraph({ type: 'bookmarkStart', attrs: { name: 'Alvo', bid: '1' } }, text('x'), {
          type: 'bookmarkEnd',
          attrs: { bid: '1' },
        }),
        paragraph(field(' PAGEREF Alvo \\h ', '1')),
      ],
    })
    expect(updateFields(editor, contextWith({})).changed).toBe(0)

    // Pagination changes afterwards (the target would move to sheet 2), and the pass must not
    // rewrite the field F9 left as it was.
    settlePageFields(editor, contextWith({ pageStarts: [{ blockIndex: 0 }] }))
    expect(fields()).toEqual(['1'])
  })
})

describe('referência cruzada a uma nota (M11)', () => {
  const footnote = (body: string): unknown => ({
    type: 'noteRef',
    attrs: { kind: 'footnote', nid: null, mark: null },
    content: [paragraph(text(body))],
  })

  it('acha a nota pelo marcador e mostra o número dela', () => {
    const plain = ProseMirrorNode.fromJSON(schema, {
      type: 'doc',
      content: [paragraph(text('A'), footnote('Um.'), text('B'), footnote('Dois.'))],
    })
    const refs: number[] = []
    plain.descendants((node, pos) => {
      if (node.type.name !== 'noteRef') return true
      refs.push(pos)
      return false
    })
    expect(noteNumberIn(plain, ['1', '2'], refs[1]!, refs[1]! + 1)).toBe('2')
    expect(noteNumberIn(plain, ['1', '2'], 0, 1)).toBeNull()
    expect(
      crossReferenceTargets(plain, BUILTIN_STYLES, { type: 'note', kind: 'footnote' }, ['1', '2']).map(
        (target) => target.text,
      ),
    ).toEqual(['1 Um.', '2 Dois.'])
  })

  it('insere o NOTEREF com o marcador na referência, e o F9 o renumera', () => {
    const { editor, fields } = fakeEditor({
      type: 'doc',
      content: [paragraph(text('A'), footnote('Um.'), text('B'), footnote('Dois.')), paragraph(text('Ver '))],
    })
    const [target] = crossReferenceTargets(editor.state.doc, BUILTIN_STYLES, {
      type: 'note',
      kind: 'footnote',
    }).slice(1)
    const end = editor.state.doc.content.size - 1
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end)))
    expect(
      insertCrossReference(editor, contextWith({}), {
        kind: { type: 'note', kind: 'footnote' },
        key: target!.key,
        show: 'number',
        link: true,
      }),
    ).toBe(true)
    expect(fields()).toEqual(['2'])
    let instr = ''
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'field') instr = String(node.attrs['instr'])
      return true
    })
    expect(instr).toMatch(/^ NOTEREF _Ref\d+ \\h $/)
    const host = editor.state.doc.child(0)
    const kinds: string[] = []
    host.forEach((child) => kinds.push(child.type.name))
    expect(kinds.slice(-3)).toEqual(['bookmarkStart', 'noteRef', 'bookmarkEnd'])

    // A new note before it: F9 now cites 3.
    editor.view.dispatch(editor.state.tr.insert(1, ProseMirrorNode.fromJSON(schema, footnote('Zero.'))))
    updateFields(editor, contextWith({}))
    expect(fields()).toEqual(['3'])
  })
})
