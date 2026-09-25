import { describe, expect, it } from 'vitest'
import { Schema } from '@tiptap/pm/model'
import { slicePageBlocks } from './print-source.js'

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'text*' },
    text: {},
    table: { group: 'block', content: 'tableRow+' },
    tableRow: { content: 'tableCell+' },
    tableCell: { content: 'paragraph+' },
    bulletList: { group: 'block', content: 'listItem+' },
    orderedList: { group: 'block', content: 'listItem+', attrs: { start: { default: 1 } } },
    listItem: { content: 'paragraph+' },
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
})
