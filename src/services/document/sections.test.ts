import { describe, expect, it } from 'vitest'
import { plainBand } from './band.js'
import { DEFAULT_PAGE_SETUP, type SectionSetup } from './model.js'
import {
  blockSections,
  effectiveSections,
  parityOf,
  sectionBreakIn,
  sheetSetups,
  startsNewSheet,
} from './sections.js'

const header = plainBand('Cabeçalho da primeira')!
const first: SectionSetup = { ...DEFAULT_PAGE_SETUP, id: 's1', headerBand: header }
const second: SectionSetup = {
  ...DEFAULT_PAGE_SETUP,
  id: 's2',
  orientation: 'landscape',
  start: 'continuous',
}

describe('seções', () => {
  it('a faixa que a seção não declara vem da anterior; a declarada vazia não herda', () => {
    const empty = { ...header, left: [], center: [], right: [] }
    const [a, b, c] = effectiveSections({ ...DEFAULT_PAGE_SETUP, headerBand: empty }, [first, second])
    expect(a!.headerBand).toBe(header)
    expect(b!.headerBand).toBe(header)
    expect(c!.headerBand).toBe(empty)
  })

  it('o bloco é da seção da próxima marca; depois da última, da seção do corpo', () => {
    expect(blockSections([null, 's1', null, 's2', null], [first, second])).toEqual([0, 0, 1, 1, 2])
    // Marca de id desconhecido (parágrafo colado de outro documento) não é marca.
    expect(blockSections(['x9', null], [first])).toEqual([1, 1])
  })

  it('a marca num item de lista fecha a seção com a lista', () => {
    const paragraph = { attrs: { sectionBreak: 's1' }, isTextblock: true }
    const list = {
      attrs: {},
      descendants: (visit: (node: typeof paragraph) => boolean | void) => void visit(paragraph),
    }
    expect(sectionBreakIn(list)).toBe('s1')
    expect(sectionBreakIn({ attrs: { sectionBreak: 's2' }, isTextblock: true })).toBe('s2')
  })

  it('a contínua só abre folha quando o papel muda; par e ímpar pedem paridade', () => {
    expect(startsNewSheet({ ...DEFAULT_PAGE_SETUP, start: 'continuous' }, DEFAULT_PAGE_SETUP)).toBe(false)
    expect(startsNewSheet(second, DEFAULT_PAGE_SETUP)).toBe(true)
    expect(startsNewSheet(DEFAULT_PAGE_SETUP, DEFAULT_PAGE_SETUP)).toBe(true)
    expect(startsNewSheet(DEFAULT_PAGE_SETUP, undefined)).toBe(false)
    expect(parityOf({ ...DEFAULT_PAGE_SETUP, start: 'oddPage' })).toBe('odd')
    expect(parityOf({ ...DEFAULT_PAGE_SETUP, start: 'evenPage' })).toBe('even')
  })

  it('cada folha conta dentro da sua seção, e a em branco não é capa', () => {
    const titled = { ...DEFAULT_PAGE_SETUP, titlePage: true }
    const setups = sheetSetups(
      [DEFAULT_PAGE_SETUP, titled],
      [
        { section: 0, number: 1, first: true, blank: false },
        { section: 0, number: 2, first: false, blank: false },
        { section: 1, number: 3, first: false, blank: true },
        { section: 1, number: 4, first: true, blank: false },
      ],
    )
    expect(setups.map((sheet) => [sheet.inSection, sheet.page.pageNumberStart])).toEqual([
      [1, 1],
      [2, 1],
      [1, 3],
      [1, 4],
    ])
    expect(setups[2]!.page.titlePage).toBe(false)
    expect(setups[3]!.page.titlePage).toBe(true)
  })
})
