import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { splitBlock } from '@tiptap/pm/commands'
import { EditorState, TextSelection, type Transaction } from '@tiptap/pm/state'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { buildEditorExtensions } from '../editor-extensions.js'
import { uniqueOids } from './block-identity.js'

/**
 * A identidade do bloco é única no documento.
 *
 * O `oid` é o que decide o que **não** reescrever ao salvar: o bloco que o
 * carrega volta para o `.docx` como o XML original. Dois blocos com o mesmo
 * `oid` é o pior caso possível — o gravador preserva o XML do primeiro e
 * regenera o segundo, e o segundo perde o que o editor não sabe reproduzir.
 *
 * Era o que o Enter fazia: dividir um parágrafo entregava os dois lados com o
 * `oid` do original, e o lado **não editado** voltava regenerado — foi assim que
 * o marcador de um parágrafo dividido desapareceu do arquivo.
 *
 * Sem `Editor` do Tiptap, que precisa de DOM: o schema sai das extensões reais e
 * o Enter é o `splitBlock` do ProseMirror, que é exatamente o comando que a tecla
 * dispara.
 */

const schema = getSchema(buildEditorExtensions(() => {}))

function stateOf(doc: unknown): EditorState {
  return EditorState.create({ doc: schema.nodeFromJSON(doc), plugins: [uniqueOids()] })
}

/** Aplica um comando como o editor o aplica: o despacho passa pelo `apply`. */
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

    // O cursor entre "Antes" e " e depois": 1 é o começo do parágrafo.
    const divided = run(state, 1 + 'Antes'.length, splitBlock)

    expect(divided.doc.childCount).toBe(2)
    expect(oidsOf(divided.doc)).toEqual(['b1', null])
  })

  it('a identidade repetida por colagem também sai', () => {
    // A gravação já previa o `oid` repetido — preserva o XML da primeira
    // ocorrência e regenera as outras. Aqui a regra é a mesma, um passo antes:
    // a segunda ocorrência deixa de afirmar uma identidade que não é dela.
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
    // O plugin devolve transação só quando há o que corrigir: uma transação
    // apendada a cada tecla digitada suja o histórico de desfazer.
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
