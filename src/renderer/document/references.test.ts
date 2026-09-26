import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { BUILTIN_STYLES } from '@services/document/styles.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { EditorState, type Transaction } from '@tiptap/pm/state'
import type { Editor } from '@tiptap/react'
import { DEFAULT_PAGE_SETUP } from '@services/document/model.js'
import {
  captionLabels,
  crossReferenceTargets,
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
    // A segunda folha abre no terceiro bloco; a terceira, no meio do quarto.
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

/** Um editor de mentira: o estado e o `dispatch`, que é tudo o que as funções usam. */
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

    // A paginação muda depois — o alvo passaria para a folha 2 —, e o passe não
    // pode reescrever o campo que o F9 deixou como estava.
    settlePageFields(editor, contextWith({ pageStarts: [{ blockIndex: 0 }] }))
    expect(fields()).toEqual(['1'])
  })
})
