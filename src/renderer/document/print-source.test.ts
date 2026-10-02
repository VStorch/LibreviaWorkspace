import { describe, expect, it } from 'vitest'
import { Schema, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { RevisionView } from '@shared/types.js'
import { blocksForView, slicePageBlocks } from './print-source.js'

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'text*', attrs: { markRevision: { default: null } } },
    text: {},
    table: { group: 'block', content: 'tableRow+' },
    tableRow: { content: '(tableCell | tableHeader)+', attrs: { rowRevision: { default: null } } },
    tableCell: { content: 'paragraph+' },
    tableHeader: { content: 'paragraph+' },
    bulletList: { group: 'block', content: 'listItem+' },
    orderedList: { group: 'block', content: 'listItem+', attrs: { start: { default: 1 } } },
    listItem: { content: 'paragraph+' },
  },
  marks: {
    insertion: { attrs: { author: { default: null } }, inclusive: false },
    deletion: { attrs: { author: { default: null } }, inclusive: false },
  },
})
const paragraph = (text: string) => schema.node('paragraph', null, schema.text(text))
const row = (text: string) => schema.node('tableRow', null, schema.node('tableCell', null, paragraph(text)))
const item = (text: string) => schema.node('listItem', null, paragraph(text))

describe('recorte das páginas para impressão', () => {
  it('recorta uma tabela em três folhas sem repetir ou perder linhas', () => {
    const table = schema.node('table', null, ['A', 'B', 'C', 'D', 'E'].map(row))
    const blocks = [paragraph('Antes'), table, paragraph('Depois')]
    const cuts = [
      { blockIndex: 0 },
      { blockIndex: 1, childIndex: 2 },
      { blockIndex: 1, childIndex: 4 },
      { blockIndex: 3 },
    ]
    const pages = cuts.slice(0, -1).map((start, index) => slicePageBlocks(blocks, start, cuts[index + 1]!))
    expect(pages.map((page) => page.map((node) => node.textContent))).toEqual([
      ['Antes', 'AB'],
      ['CD'],
      ['E', 'Depois'],
    ])
    expect(pages[1]![0]!.type.name).toBe('table')
  })

  it('lista numerada continua de onde a folha anterior terminou', () => {
    const list = schema.node('orderedList', { start: 7 }, ['A', 'B', 'C'].map(item))
    const page = slicePageBlocks([list], { blockIndex: 0, childIndex: 2 }, { blockIndex: 1 })
    expect(page[0]!.attrs.start).toBe(9)
    expect(page[0]!.textContent).toBe('C')
  })

  it('preserva a estrutura da lista de marcadores', () => {
    const list = schema.node('bulletList', null, ['A', 'B', 'C'].map(item))
    const page = slicePageBlocks([list], { blockIndex: 0 }, { blockIndex: 0, childIndex: 1 })
    expect(page[0]!.type.name).toBe('bulletList')
    expect(page[0]!.textContent).toBe('A')
  })

  it('mantém blocos inteiros e a folha vazia de uma quebra explícita final', () => {
    const block = paragraph('Inteiro')
    expect(slicePageBlocks([block], { blockIndex: 0 }, { blockIndex: 1 })).toEqual([block])
    expect(slicePageBlocks([block], { blockIndex: 1 }, { blockIndex: 1 })).toEqual([])
  })

  it('parágrafo cortado entre linhas sai em dois pedaços que somam o original', () => {
    const block = paragraph('Primeira linha. Segunda linha.')
    const blocks = [paragraph('Antes'), block]
    const cut = { blockIndex: 1, offset: 16 }
    const top = slicePageBlocks(blocks, { blockIndex: 0 }, cut)
    const bottom = slicePageBlocks(blocks, cut, { blockIndex: 2 })
    expect(top.map((node) => node.textContent)).toEqual(['Antes', 'Primeira linha. '])
    expect(bottom.map((node) => node.textContent)).toEqual(['Segunda linha.'])
  })

  it('parágrafo de três folhas: o pedaço do meio é só o miolo', () => {
    const block = paragraph('aaaabbbbcccc')
    const page = slicePageBlocks([block], { blockIndex: 0, offset: 4 }, { blockIndex: 0, offset: 8 })
    expect(page.map((node) => node.textContent)).toEqual(['bbbb'])
  })

  it('a folha em que a tabela continua abre com as linhas de cabeçalho', () => {
    const header = schema.node('tableRow', null, schema.node('tableHeader', null, paragraph('Cab')))
    const table = schema.node('table', null, [header, ...['A', 'B', 'C'].map(row)])
    const page = slicePageBlocks(
      [table],
      { blockIndex: 0, childIndex: 2, repeatHeader: true },
      { blockIndex: 1 },
    )
    expect(page[0]!.textContent).toBe('CabBC')
    const plain = slicePageBlocks([table], { blockIndex: 0, childIndex: 2 }, { blockIndex: 1 })
    expect(plain[0]!.textContent).toBe('BC')
  })
})

describe('impressão conforme Revisão → Mostrar', () => {
  const ins = schema.marks['insertion']!.create({ author: 'Ana' })
  const del = schema.marks['deletion']!.create({ author: 'Ana' })
  const revised = schema.node('paragraph', { markRevision: { kind: 'del', author: 'Ana' } }, [
    schema.text('Era '),
    schema.text('velho', [del]),
    schema.text('novo', [ins]),
  ])
  const insertedBlock = schema.node(
    'paragraph',
    { markRevision: { kind: 'ins' } },
    schema.text('Todo novo', [ins]),
  )
  const deletedBlock = schema.node(
    'paragraph',
    { markRevision: { kind: 'del' } },
    schema.text('Todo velho', [del]),
  )
  const table = schema.node('table', null, [
    row('Fica'),
    schema.node(
      'tableRow',
      { rowRevision: { kind: 'ins' } },
      schema.node('tableCell', null, paragraph('Nova')),
    ),
    schema.node(
      'tableRow',
      { rowRevision: { kind: 'del' } },
      schema.node('tableCell', null, paragraph('Velha')),
    ),
  ])
  const blocks = [revised, insertedBlock, deletedBlock, table]
  const hasBlockRevision = (node: ProseMirrorNode): boolean =>
    (node.attrs['markRevision'] ?? null) !== null || (node.attrs['rowRevision'] ?? null) !== null
  const marksOf = (nodes: readonly ProseMirrorNode[]): string[] => {
    const names = new Set<string>()
    for (const node of nodes) {
      node.descendants((child) => {
        for (const mark of child.marks) names.add(mark.type.name)
        if (hasBlockRevision(child)) names.add('bloco')
      })
      if (hasBlockRevision(node)) names.add('bloco')
    }
    return [...names].sort()
  }

  it('marcação completa: tudo como está, com as marcas', () => {
    const shown = blocksForView(blocks, RevisionView.All)
    expect(shown).toEqual(blocks)
    expect(marksOf(shown)).toEqual(['bloco', 'deletion', 'insertion'])
  })

  it.each([RevisionView.Simple, RevisionView.None])('%s: o texto final, sem excluído nem marca', (view) => {
    const shown = blocksForView(blocks, view)
    expect(shown.map((node) => node.textContent)).toEqual(['Era novo', 'Todo novo', 'FicaNova'])
    expect(marksOf(shown)).toEqual([])
  })

  it('original: o texto de antes, sem inserido nem marca', () => {
    const shown = blocksForView(blocks, RevisionView.Original)
    expect(shown.map((node) => node.textContent)).toEqual(['Era velho', 'Todo velho', 'FicaVelha'])
    expect(marksOf(shown)).toEqual([])
  })

  it('a célula que perdeu o texto continua, vazia', () => {
    const cell = schema.node('tableCell', null, schema.node('paragraph', null, schema.text('x', [ins])))
    const shown = blocksForView(
      [schema.node('table', null, schema.node('tableRow', null, cell))],
      RevisionView.Original,
    )
    expect(shown[0]!.firstChild!.childCount).toBe(1)
    expect(shown[0]!.textContent).toBe('')
  })
})
