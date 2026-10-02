import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { Fragment, Node as ProseMirrorNode, Slice } from '@tiptap/pm/model'
import { EditorState, TextSelection, type Transaction } from '@tiptap/pm/state'
import { history, undo, undoDepth } from '@tiptap/pm/history'
import { buildEditorExtensions } from '../editor-extensions.js'
import { revisionChangesOf, settleAllChanges, textWithoutDeletions } from './track-changes.js'
import {
  SKIP_TRACKING,
  joinHistoryGroup,
  type TrackGroup,
  revisionDate,
  shouldTrack,
  stripDeleted,
  stripRevisions,
  trackTransaction,
  wordRangeAt,
} from './track-input.js'

const schema = getSchema(buildEditorExtensions(() => {}))
const NOW = new Date('2026-10-01T12:34:56Z')
const DATE = '2026-10-01T12:34:00Z'

const mark = (type: 'insertion' | 'deletion', author: string): Record<string, unknown> => ({
  type,
  attrs: { author, date: '2026-01-01T00:00:00Z', rid: '7' },
})

function paragraph(...content: Record<string, unknown>[]): Record<string, unknown> {
  return { type: 'paragraph', content }
}

function text(value: string, ...marks: Record<string, unknown>[]): Record<string, unknown> {
  return marks.length === 0 ? { type: 'text', text: value } : { type: 'text', text: value, marks }
}

function stateOf(blocks: Record<string, unknown>[], cursor?: number, withHistory = false): EditorState {
  const doc = ProseMirrorNode.fromJSON(schema, { type: 'doc', content: blocks })
  const state = EditorState.create({ doc, ...(withHistory ? { plugins: [history()] } : {}) })
  return cursor === undefined ? state : state.apply(state.tr.setSelection(TextSelection.create(doc, cursor)))
}

/** Aplica a transação com o controle ligado, como Ana. */
function track(state: EditorState, build: (tr: Transaction) => Transaction, options = {}): EditorState {
  const tracked = trackTransaction(build(state.tr), state, 'Ana', NOW, options)
  const next = state.apply(tracked)
  next.doc.check()
  return next
}

/** O texto de cada parágrafo, com o excluído entre colchetes e o inserido entre chaves. */
function shown(doc: ProseMirrorNode): string[] {
  const blocks: string[] = []
  doc.descendants((node) => {
    if (!node.isTextblock) return true
    let line = ''
    node.forEach((child) => {
      const value = child.text ?? ''
      const deleted = child.marks.some((m) => m.type.name === 'deletion')
      const inserted = child.marks.some((m) => m.type.name === 'insertion')
      line += deleted && inserted ? `{[${value}]}` : deleted ? `[${value}]` : inserted ? `{${value}}` : value
    })
    const revision = node.attrs['markRevision'] as { kind: string } | null
    blocks.push(revision === null ? line : `${line}¶${revision.kind}`)
    return false
  })
  return blocks
}

function marksAt(doc: ProseMirrorNode, pos: number): Record<string, unknown>[] {
  return (doc.nodeAt(pos)?.marks ?? []).map((m) => ({ type: m.type.name, ...m.attrs }))
}

describe('controle do que se digita', () => {
  it('o que se digita vira inserção do autor, e o mesmo minuto é um trecho só', () => {
    let state = stateOf([paragraph(text('abc'))], 4)
    state = track(state, (tr) => tr.insertText('x'))
    state = track(state, (tr) => tr.insertText('y'))
    expect(shown(state.doc)).toEqual(['abc{xy}'])
    expect(marksAt(state.doc, 4)).toEqual([
      { type: 'insertion', author: 'Ana', date: DATE, rid: null, move: null, moveName: null },
    ])
    expect(revisionChangesOf(state.doc)).toHaveLength(1)
    expect(state.selection.from).toBe(6)
  })

  it('digitar no meio da inserção de outro autor não herda a marca dele', () => {
    let state = stateOf([paragraph(text('ab', mark('insertion', 'Bruno')))], 2)
    state = track(state, (tr) => tr.insertText('x'))
    expect(marksAt(state.doc, 2)[0]).toMatchObject({ type: 'insertion', author: 'Ana' })
  })

  it('o Backspace guarda o caractere excluído e deixa o cursor antes dele', () => {
    let state = stateOf([paragraph(text('abc'))], 4)
    state = track(state, (tr) => tr.delete(3, 4))
    expect(shown(state.doc)).toEqual(['ab[c]'])
    expect(state.selection.from).toBe(3)
    // O seguinte exclui o anterior, e o trecho se junta.
    state = track(state, (tr) => tr.delete(2, 3))
    expect(shown(state.doc)).toEqual(['a[bc]'])
    expect(state.selection.from).toBe(2)
  })

  it('o Delete guarda o caractere e deixa o cursor depois dele', () => {
    let state = stateOf([paragraph(text('abc'))], 2)
    state = track(state, (tr) => tr.delete(2, 3))
    expect(shown(state.doc)).toEqual(['a[b]c'])
    expect(state.selection.from).toBe(3)
  })

  it('apagar a própria inserção apaga de verdade', () => {
    let state = stateOf([paragraph(text('ab'), text('xy', mark('insertion', 'Ana')), text('c'))], 5)
    state = track(state, (tr) => tr.delete(4, 5))
    expect(shown(state.doc)).toEqual(['ab{x}c'])
    expect(state.selection.from).toBe(4)
  })

  it('apagar a inserção de outro autor a deixa inserida e excluída', () => {
    let state = stateOf([paragraph(text('a'), text('xy', mark('insertion', 'Bruno')))], 4)
    state = track(state, (tr) => tr.delete(2, 4))
    expect(shown(state.doc)).toEqual(['a{[xy]}'])
  })

  it('o que já é exclusão fica como está: o cursor só passa por cima', () => {
    let state = stateOf([paragraph(text('a'), text('b', mark('deletion', 'Bruno')))], 3)
    const before = state.doc
    state = track(state, (tr) => tr.delete(2, 3))
    expect(state.doc.eq(before)).toBe(true)
    expect(state.selection.from).toBe(2)
  })

  it('digitar sobre a seleção exclui o selecionado e insere depois dele', () => {
    let state = stateOf([paragraph(text('um dois três'))])
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 4, 8)))
    state = track(state, (tr) => tr.insertText('2'))
    expect(shown(state.doc)).toEqual(['um [dois]{2} três'])
    expect(state.selection.from).toBe(9)
  })

  it('apagar entre parágrafos mantém os dois e exclui a marca do primeiro', () => {
    let state = stateOf([paragraph(text('abc')), paragraph(text('def'))])
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3, 8)))
    state = track(state, (tr) => tr.deleteSelection())
    expect(shown(state.doc)).toEqual(['ab[c]¶del', '[de]f'])
    const accept = state.tr
    settleAllChanges(accept, true)
    expect(shown(state.apply(accept).doc)).toEqual(['abf'])
  })

  it('o Backspace no começo do parágrafo exclui a marca do de cima', () => {
    let state = stateOf([paragraph(text('ab')), paragraph(text('cd'))], 5)
    state = track(state, (tr) => tr.join(4))
    expect(shown(state.doc)).toEqual(['ab¶del', 'cd'])
    expect(state.selection.from).toBe(3)
  })

  it('o Enter insere a marca de parágrafo; o Backspace logo depois a tira de verdade', () => {
    let state = stateOf([paragraph(text('abcd'))], 3)
    state = track(state, (tr) => tr.split(3))
    expect(shown(state.doc)).toEqual(['ab¶ins', 'cd'])
    expect(state.selection.from).toBe(5)
    state = track(state, (tr) => tr.join(4))
    expect(shown(state.doc)).toEqual(['abcd'])
  })

  it('o Backspace no parágrafo vazio criado pelo Enter também o desfaz de verdade', () => {
    let state = stateOf([paragraph(text('ab'))], 3)
    state = track(state, (tr) => tr.split(3))
    expect(shown(state.doc)).toEqual(['ab¶ins', ''])
    state = track(state, (tr) => tr.delete(4, 6))
    expect(shown(state.doc)).toEqual(['ab'])
  })

  it('a colagem perde as revisões de fora e entra como inserção', () => {
    const pasted = new Slice(
      Fragment.from(schema.text('cola', [schema.marks['deletion']!.create({ author: 'Bruno' })])),
      0,
      0,
    )
    let state = stateOf([paragraph(text('ab'))], 2)
    state = track(state, (tr) => tr.replaceSelection(stripRevisions(pasted)))
    expect(shown(state.doc)).toEqual(['a{cola}b'])
  })

  it('copiar deixa o excluído de fora', () => {
    const state = stateOf([paragraph(text('a'), text('xx', mark('deletion', 'Bruno')), text('b'))])
    const slice = stripDeleted(state.doc.slice(1, 5))
    expect(slice.content.textBetween(0, slice.content.size)).toBe('ab')
  })

  it('um desfazer devolve o documento de antes', () => {
    let state = stateOf([paragraph(text('abc'))], 2, true)
    const before = state.doc
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 3)))
    state = track(state, (tr) => tr.insertText('zz'))
    expect(shown(state.doc)).toEqual(['[ab]{zz}c'])
    let undone = state
    undo(state, (tr) => (undone = state.apply(tr)))
    expect(undone.doc.eq(before)).toBe(true)
  })

  it('Backspaces seguidos, Delete seguidos, e Enter com o que se digita depois: um desfazer cada', () => {
    const initial = stateOf([paragraph(text('abcdef'))], 4, true)
    let state = initial
    let group: TrackGroup | null = null
    let time = 1_000
    // Como o `dispatchTransaction` faz: reescreve e junta ao grupo.
    const edit = (build: (tr: Transaction) => Transaction): void => {
      const original = build(state.tr).setTime((time += 100))
      const tracked = trackTransaction(original, state, 'Ana', NOW)
      group = joinHistoryGroup(original, tracked, group)
      state = state.apply(tracked)
    }
    edit((tr) => tr.delete(3, 4))
    edit((tr) => tr.delete(2, 3))
    edit((tr) => tr.delete(1, 2))
    expect(shown(state.doc)).toEqual(['[abc]def'])
    expect(undoDepth(state)).toBe(1)

    // O Delete, que deixa o cursor depois do excluído.
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 4)))
    time += 1_000
    edit((tr) => tr.delete(4, 5))
    edit((tr) => tr.delete(5, 6))
    expect(shown(state.doc)).toEqual(['[abcde]f'])
    expect(undoDepth(state)).toBe(2)

    // O Enter e o texto do parágrafo novo.
    time += 1_000
    edit((tr) => tr.split(state.selection.from))
    edit((tr) => tr.insertText('x'))
    edit((tr) => tr.insertText('y'))
    expect(undoDepth(state)).toBe(3)

    // E desfazer devolve cada grupo exatamente.
    for (let step = 0; step < 3; step++) undo(state, (tr) => (state = state.apply(tr)))
    expect(state.doc.eq(initial.doc)).toBe(true)
  })

  it('depois de meio segundo, ou longe do cursor, começa outro desfazer', () => {
    let state = stateOf([paragraph(text('abcdef'))], 4, true)
    const tracked1 = trackTransaction(state.tr.delete(3, 4).setTime(1_000), state, 'Ana', NOW)
    let group = joinHistoryGroup(state.tr.delete(3, 4).setTime(1_000), tracked1, null)
    state = state.apply(tracked1)
    const late = state.tr.delete(2, 3).setTime(2_000)
    const tracked2 = trackTransaction(late, state, 'Ana', NOW)
    group = joinHistoryGroup(late, tracked2, group)
    state = state.apply(tracked2)
    const far = state.tr.delete(6, 7).setTime(2_100)
    const tracked3 = trackTransaction(far, state, 'Ana', NOW)
    joinHistoryGroup(far, tracked3, group)
    state = state.apply(tracked3)
    expect(undoDepth(state)).toBe(3)
  })

  it('na composição do IME o que sai sai de verdade, e o que entra é marcado', () => {
    let state = stateOf([paragraph(text('abc'))], 3)
    state = track(state, (tr) => tr.insertText('か', 2, 3), { composing: true })
    expect(shown(state.doc)).toEqual(['a{か}c'])
  })

  it('substituir tudo: cada ocorrência excluída e a nova inserida depois', () => {
    const blocks = [paragraph(text('gato e gato')), paragraph(text('outro gato'))]
    const replaceAll = (tr: Transaction): Transaction => {
      for (const from of [20, 8, 1]) tr.insertText('cão', from, from + 4)
      return tr
    }
    const plain = stateOf(blocks).apply(replaceAll(stateOf(blocks).tr))
    const state = track(stateOf(blocks), replaceAll)
    expect(shown(state.doc)).toEqual(['[gato]{cão} e [gato]{cão}', 'outro [gato]{cão}'])
    expect(textWithoutDeletions(state.doc, '\n', '')).toBe(
      plain.doc.textBetween(0, plain.doc.content.size, '\n'),
    )
  })

  it('a linha de tabela apagada fica excluída; a inserida, marcada', () => {
    const cell = (value: string): Record<string, unknown> => ({
      type: 'tableCell',
      content: [paragraph(text(value))],
    })
    const row = (value: string): Record<string, unknown> => ({ type: 'tableRow', content: [cell(value)] })
    let state = stateOf([{ type: 'table', content: [row('a'), row('b')] }, paragraph()])
    const first = state.doc.firstChild!.firstChild!
    state = track(state, (tr) => tr.delete(1, 1 + first.nodeSize))
    expect(state.doc.firstChild!.childCount).toBe(2)
    expect(state.doc.firstChild!.firstChild!.attrs['rowRevision']).toMatchObject({
      kind: 'del',
      author: 'Ana',
    })

    const added = schema.nodes['tableRow']!.create(null, schema.nodes['tableCell']!.createAndFill()!)
    const end = 1 + state.doc.firstChild!.content.size
    state = track(state, (tr) => tr.insert(end, added))
    expect(state.doc.firstChild!.lastChild!.attrs['rowRevision']).toMatchObject({
      kind: 'ins',
      author: 'Ana',
    })
    // A linha inserida pelo próprio autor sai de verdade.
    const last = state.doc.firstChild!.lastChild!
    state = track(state, (tr) => tr.delete(end, end + last.nodeSize))
    expect(state.doc.firstChild!.childCount).toBe(2)
  })

  it('desfazer, aceitar e o que não entra no histórico passam sem controle', () => {
    const state = stateOf([paragraph(text('abc'))])
    expect(shouldTrack(state.tr.insertText('x', 1))).toBe(true)
    expect(shouldTrack(state.tr)).toBe(false)
    expect(shouldTrack(state.tr.insertText('x', 1).setMeta(SKIP_TRACKING, true))).toBe(false)
    expect(shouldTrack(state.tr.insertText('x', 1).setMeta('addToHistory', false))).toBe(false)
    expect(shouldTrack(state.tr.insertText('x', 1).setMeta('history$', { redo: false }))).toBe(false)
  })

  it('a ponta de comentário entra e sai sem virar revisão', () => {
    let state = stateOf([paragraph(text('abc'))])
    const anchor = schema.nodes['commentStart']!.create({ cid: '1' })
    state = track(state, (tr) => tr.insert(2, anchor))
    expect(revisionChangesOf(state.doc)).toHaveLength(0)
    state = track(state, (tr) => tr.delete(2, 3))
    expect(state.doc.textContent).toBe('abc')
    expect(revisionChangesOf(state.doc)).toHaveLength(0)
  })

  it('a data tem precisão de minuto', () => {
    expect(revisionDate(NOW)).toBe(DATE)
  })
})

/** Um gerador pequeno e determinístico, para o teste aleatório repetir igual. */
function random(seed: number): () => number {
  let value = seed
  return () => {
    value = (value * 1103515245 + 12345) % 2147483648
    return value / 2147483648
  }
}

function textPositions(doc: ProseMirrorNode): number[] {
  const positions: number[] = []
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    for (let offset = 0; offset <= node.content.size; offset++) positions.push(pos + 1 + offset)
    return false
  })
  return positions
}

/**
 * O texto depois de aceitar (ou rejeitar) tudo. Aceito, sem as quebras de
 * parágrafo: a edição sem controle junta blocos com os atributos do de cima, e a
 * marca de parágrafo que ela perde não é a referência certa.
 */
function settled(state: EditorState, accept: boolean): string {
  const tr = state.tr
  settleAllChanges(tr, accept)
  const doc = state.apply(tr).doc
  return doc.textBetween(0, doc.content.size, accept ? '' : '\n')
}

describe('controle do que se digita — aleatório', () => {
  it('o documento segue válido; aceitar dá o texto sem controle, rejeitar dá o de antes', () => {
    const next = random(42)
    const pick = <T>(items: readonly T[]): T => items[Math.floor(next() * items.length)]!
    const words = ['a', 'bc', 'def', ' ', 'xyz']

    /** Uma edição de um ou dois passos em posições de texto ao acaso. */
    const edit = (tr: Transaction): Transaction => {
      const steps = 1 + Math.floor(next() * 2)
      for (let i = 0; i < steps; i++) {
        const positions = textPositions(tr.doc)
        const a = pick(positions)
        const b = pick(positions)
        const [from, to] = a <= b ? [a, b] : [b, a]
        const kind = next()
        // Texto sem marca: o `insertText` herdaria a exclusão de onde cai, e a
        // edição sem controle deixaria de ser a referência.
        if (kind < 0.35) tr.replaceWith(from, to, schema.text(pick(words)))
        else if (kind < 0.7 && to > from) tr.delete(from, to)
        else if (kind < 0.8) tr.split(from)
        else if (kind < 0.9) {
          // A colagem de dois parágrafos no lugar do trecho.
          const pasted = [pick(words), pick(words)].map((word) =>
            schema.nodes['paragraph']!.create(null, schema.text(word)),
          )
          tr.replace(from, to, new Slice(Fragment.from(pasted), 1, 1))
        } else tr.insert(from, schema.text(pick(words)))
      }
      return tr
    }

    for (let round = 0; round < 200; round++) {
      let state = stateOf([
        paragraph(text('primeiro '), text('velho', mark('insertion', 'Bruno')), text(' fim')),
        paragraph(text('segundo '), text('sumido', mark('deletion', 'Bruno'))),
        paragraph(text('terceiro')),
      ])
      const original = settled(state, false)

      for (let step = 0; step < 4; step++) {
        const tr = edit(state.tr)
        const plain = state.apply(tr)
        const tracked = state.apply(trackTransaction(tr, state, 'Ana', NOW))
        tracked.doc.check()
        expect(settled(tracked, true)).toBe(settled(plain, true))
        expect(settled(tracked, false)).toBe(original)
        state = tracked
      }
    }
  })
})

describe('Ctrl+Backspace e Ctrl+Delete', () => {
  it('a palavra do texto que se vê, passando por cima do já excluído', () => {
    const block = stateOf([
      paragraph(text('one '), text('two', mark('deletion', 'Bia')), text(' three')),
    ]).doc.child(0)
    // "one two three": o cursor no fim apaga "three"; antes de "three", o espaço
    // e "two" (excluído, transparente) e chega a "one".
    expect(wordRangeAt(block, 13, true)).toEqual([8, 13])
    expect(wordRangeAt(block, 8, true)).toEqual([0, 8])
    expect(wordRangeAt(block, 0, false)).toEqual([0, 8])
    expect(wordRangeAt(block, 0, true)).toBeNull()
  })
})
