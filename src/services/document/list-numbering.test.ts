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

/** "1. / 1.1." with the second level in letters, as manuals often have. */
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
  it('writes each Word format', () => {
    expect(formatNumber(4, 'decimal')).toBe('4')
    expect(formatNumber(4, 'decimalZero')).toBe('04')
    expect(formatNumber(4, 'lowerRoman')).toBe('iv')
    expect(formatNumber(1994, 'upperRoman')).toBe('MCMXCIV')
    expect(formatNumber(2, 'upperLetter')).toBe('B')
    expect(formatNumber(4, 'bullet')).toBe('')
  })

  it('after z comes aa, as in Word, not ab', () => {
    expect(formatNumber(26, 'lowerLetter')).toBe('z')
    expect(formatNumber(27, 'lowerLetter')).toBe('aa')
    expect(formatNumber(28, 'lowerLetter')).toBe('bb')
  })

  it('unknown format comes out as decimal, and does not vanish', () => {
    expect(formatNumber(3, 'ordinal')).toBe('3')
  })
})

describe('numberLists', () => {
  it("a new numbered list uses Word's default levels: 1. a. i.", () => {
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

  it('composes the upper numbers through the level text', () => {
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

  it('the same definition continues the count across a paragraph', () => {
    const root = doc(
      list('orderedList', { numId: 5, numbering: outline }, item('um'), item('dois')),
      paragraph('no meio'),
      list('orderedList', { numId: 5, numbering: outline }, item('três')),
    )
    expect(labels(root)).toEqual(['1.', '2.', '3.'])
  })

  it('another w:num of the same abstract definition also continues', () => {
    const other = { ...outline }
    const root = doc(
      list('orderedList', { numId: 5, numbering: outline }, item('um')),
      paragraph('no meio'),
      list('orderedList', { numId: 6, numbering: other }, item('dois')),
    )
    expect(labels(root)).toEqual(['1.', '2.'])
  })

  it('the w:num restart counts separately, from the given value', () => {
    const restarted: NumberingDef = { ...outline, key: 'n7', overrides: { '0': 10 } }
    const root = doc(
      list('orderedList', { numId: 5, numbering: outline }, item('um'), item('dois')),
      list('orderedList', { numId: 7, numbering: restarted }, item('dez'), item('onze')),
      list('orderedList', { numId: 5, numbering: outline }, item('três')),
    )
    expect(labels(root)).toEqual(['1.', '2.', '10.', '11.', '3.'])
  })

  it('the upper item restarts the lower ones', () => {
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

  it('w:isLgl writes the upper numbers as decimal', () => {
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

  it('a bulleted list draws the level mark, and a sublist created with Tab inherits the definition', () => {
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

  it('two new lists each count their own', () => {
    const root = doc(
      list('orderedList', {}, item('um'), item('dois')),
      paragraph('no meio'),
      list('orderedList', {}, item('um de novo')),
    )
    expect(labels(root)).toEqual(['1.', '2.', '1.'])
  })

  it("the file's level holds when the list starts deep", () => {
    const root = doc(list('orderedList', { numId: 5, numbering: outline, level: 1 }, item('solto')))
    const result = numberLists(root, reader)
    expect(result.lists[0]!.level).toBe(1)
    expect(result.labels).toEqual(['1.a)'])
  })

  it("the sublist indent is relative to the outer list's", () => {
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

describe('ready-made lists', () => {
  it('all define the nine levels', () => {
    for (const preset of LIST_PRESETS) {
      expect(preset.levels).toHaveLength(9)
      expect(kindOfLevels(preset.levels)).toBe(preset.kind)
    }
  })

  it('1. 1.1. 1.1.1. composes all the upper levels', () => {
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
