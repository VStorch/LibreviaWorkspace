import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { BUILTIN_STYLES } from '@services/document/styles.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { captionLabels, crossReferenceTargets, sheetAt } from './references.js'

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
