import { describe, expect, it } from 'vitest'
import type { DocumentNode } from './model.js'
import {
  defaultLevels,
  formatNumber,
  listDrawAttrs,
  numberLists,
  type LevelDef,
  type ListTreeReader,
  type NumberingDef,
} from './list-numbering.js'
import { LIST_PRESETS, kindOfLevels } from './list-presets.js'

const reader: ListTreeReader<DocumentNode> = {
  typeOf: (node) => node.type,
  attrsOf: (node) => node.attrs ?? {},
  childrenOf: (node) => node.content ?? [],
}

const paragraph = (text: string): DocumentNode => ({ type: 'paragraph', content: [{ type: 'text', text }] })
const item = (text: string, ...nested: DocumentNode[]): DocumentNode => ({
  type: 'listItem',
  content: [paragraph(text), ...nested],
})
const list = (type: string, attrs: Record<string, unknown>, ...items: DocumentNode[]): DocumentNode => ({
  type,
  attrs,
  content: items,
})
const doc = (...content: DocumentNode[]): DocumentNode => ({ type: 'doc', content })

const labels = (root: DocumentNode): string[] => numberLists(root, reader).labels

/** Definição "1. / 1.1." com o segundo nível em letra, como os manuais costumam ter. */
const outline: NumberingDef = {
  key: 'a3',
  abstractId: 3,
  levels: [
    { fmt: 'decimal', text: '%1.', start: 1, indentMm: 6.35, hangingMm: 6.35 },
    { fmt: 'lowerLetter', text: '%1.%2)', start: 1, indentMm: 12.7, hangingMm: 6.35 },
    ...defaultLevels('orderedList').slice(2),
  ],
}

describe('formatNumber', () => {
  it('escreve cada formato do Word', () => {
    expect(formatNumber(4, 'decimal')).toBe('4')
    expect(formatNumber(4, 'decimalZero')).toBe('04')
    expect(formatNumber(4, 'lowerRoman')).toBe('iv')
    expect(formatNumber(1994, 'upperRoman')).toBe('MCMXCIV')
    expect(formatNumber(2, 'upperLetter')).toBe('B')
    expect(formatNumber(4, 'bullet')).toBe('')
  })

  it('depois do z vem aa, como no Word, e não ab', () => {
    expect(formatNumber(26, 'lowerLetter')).toBe('z')
    expect(formatNumber(27, 'lowerLetter')).toBe('aa')
    expect(formatNumber(28, 'lowerLetter')).toBe('bb')
  })

  it('formato desconhecido sai em decimal, e não some', () => {
    expect(formatNumber(3, 'ordinal')).toBe('3')
  })
})

describe('numberLists', () => {
  it('lista nova numerada usa os níveis padrão do Word: 1. a. i.', () => {
    const root = doc(
      list(
        'orderedList',
        {},
        item(
          'um',
          list('orderedList', {}, item('um-a', list('orderedList', {}, item('um-a-i'))), item('um-b')),
        ),
        item('dois'),
      ),
    )
    expect(labels(root)).toEqual(['1.', 'a.', 'i.', 'b.', '2.'])
  })

  it('compõe os números de cima pelo texto do nível', () => {
    const root = doc(
      list(
        'orderedList',
        { numId: 5, numbering: outline },
        item('um'),
        item('dois', list('orderedList', { numId: 5 }, item('dois-a'), item('dois-b'))),
      ),
    )
    expect(labels(root)).toEqual(['1.', '2.', '2.a)', '2.b)'])
  })

  it('a mesma definição continua a conta do outro lado de um parágrafo', () => {
    const root = doc(
      list('orderedList', { numId: 5, numbering: outline }, item('um'), item('dois')),
      paragraph('no meio'),
      list('orderedList', { numId: 5, numbering: outline }, item('três')),
    )
    expect(labels(root)).toEqual(['1.', '2.', '3.'])
  })

  it('outro w:num da mesma definição abstrata também continua', () => {
    const other = { ...outline }
    const root = doc(
      list('orderedList', { numId: 5, numbering: outline }, item('um')),
      paragraph('no meio'),
      list('orderedList', { numId: 6, numbering: other }, item('dois')),
    )
    expect(labels(root)).toEqual(['1.', '2.'])
  })

  it('o reinício do w:num conta à parte, a partir do valor dado', () => {
    const restarted: NumberingDef = { ...outline, key: 'n7', overrides: { '0': 10 } }
    const root = doc(
      list('orderedList', { numId: 5, numbering: outline }, item('um'), item('dois')),
      list('orderedList', { numId: 7, numbering: restarted }, item('dez'), item('onze')),
      list('orderedList', { numId: 5, numbering: outline }, item('três')),
    )
    expect(labels(root)).toEqual(['1.', '2.', '10.', '11.', '3.'])
  })

  it('o item de cima reinicia os de baixo', () => {
    const root = doc(
      list(
        'orderedList',
        { numId: 5, numbering: outline },
        item('um', list('orderedList', { numId: 5 }, item('a'), item('b'))),
        item('dois', list('orderedList', { numId: 5 }, item('a de novo'))),
      ),
    )
    expect(labels(root)).toEqual(['1.', '1.a)', '1.b)', '2.', '2.a)'])
  })

  it('w:isLgl escreve os números de cima em decimal', () => {
    const levels: LevelDef[] = [
      { fmt: 'upperRoman', text: '%1.', start: 1 },
      { fmt: 'decimal', text: '%1.%2.', start: 1, legal: true },
    ]
    const root = doc(
      list(
        'orderedList',
        { numId: 1, numbering: { key: 'a1', levels } },
        item('I', list('orderedList', { numId: 1 }, item('1.1'))),
      ),
    )
    expect(labels(root)).toEqual(['I.', '1.1.'])
  })

  it('lista com marcador desenha a marca do nível, e sublista criada com Tab herda a definição', () => {
    const bullets: NumberingDef = {
      key: 'a2',
      levels: [
        { fmt: 'bullet', text: '▪', start: 1, indentMm: 12.7, hangingMm: 6.35 },
        { fmt: 'bullet', text: '➢', start: 1, indentMm: 25.4, hangingMm: 6.35 },
      ],
    }
    const root = doc(
      list(
        'bulletList',
        { numId: 2, numbering: bullets },
        item('um', list('bulletList', {}, item('dentro'))),
      ),
    )
    const result = numberLists(root, reader)
    expect(result.labels).toEqual(['▪', '➢'])
    expect(result.lists[1]!.numId).toBe(2)
  })

  it('duas listas novas contam cada uma a sua', () => {
    const root = doc(
      list('orderedList', {}, item('um'), item('dois')),
      paragraph('no meio'),
      list('orderedList', {}, item('um de novo')),
    )
    expect(labels(root)).toEqual(['1.', '2.', '1.'])
  })

  it('o nível do arquivo vale quando a lista começa fundo', () => {
    const root = doc(list('orderedList', { numId: 5, numbering: outline, level: 1 }, item('solto')))
    const result = numberLists(root, reader)
    expect(result.lists[0]!.level).toBe(1)
    expect(result.labels).toEqual(['1.a)'])
  })

  it('o recuo da sublista é relativo ao da lista de fora', () => {
    const root = doc(
      list(
        'orderedList',
        { numId: 5, numbering: outline },
        item('um', list('orderedList', { numId: 5 }, item('a'))),
      ),
    )
    const [outer, inner] = numberLists(root, reader).lists
    expect(listDrawAttrs(outer!)['style']).toContain('--lista-recuo: 6.35mm')
    expect(listDrawAttrs(inner!)['style']).toContain('--lista-recuo: 6.35mm')
    expect(inner!.indentMm).toBe(12.7)
  })
})

describe('listas prontas', () => {
  it('todas definem os nove níveis', () => {
    for (const preset of LIST_PRESETS) {
      expect(preset.levels).toHaveLength(9)
      expect(kindOfLevels(preset.levels)).toBe(preset.kind)
    }
  })

  it('1. 1.1. 1.1.1. compõe todos os níveis de cima', () => {
    const legal = LIST_PRESETS.find((preset) => preset.id === 'legal')!
    const numbering = { key: 'l', levels: legal.levels }
    const root = doc(
      list(
        'orderedList',
        { numbering },
        item('um', list('orderedList', {}, item('um-um', list('orderedList', {}, item('fundo'))))),
      ),
    )
    expect(labels(root)).toEqual(['1.', '1.1.', '1.1.1.'])
  })
})
