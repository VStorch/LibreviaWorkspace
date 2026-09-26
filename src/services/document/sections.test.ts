import { describe, expect, it } from 'vitest'
import { plainBand } from './band.js'
import { DEFAULT_PAGE_SETUP, type SectionSetup } from './model.js'
import {
  blockSections,
  effectiveSections,
  freshSectionId,
  marksOfJson,
  planSectionBreak,
  planSectionDelete,
  resolveSections,
  storeSections,
  isLinkedToPrevious,
  withBandsLinked,
  withPageSetup,
  withSectionBreak,
  withoutSection,
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

  it('a quebra nova copia a seção partida para cima, e a de baixo começa como pedido', () => {
    const page = { ...DEFAULT_PAGE_SETUP, orientation: 'landscape' as const }
    const id = freshSectionId([first])
    const split = withSectionBreak({ page, sections: [first] }, 1, id, 'oddPage')
    expect(split.sections.map((section) => [section.id, section.orientation])).toEqual([
      ['s1', 'portrait'],
      [id, 'landscape'],
    ])
    expect(split.page.start).toBe('oddPage')

    const middle = withSectionBreak({ page, sections: [first] }, 0, 'n9', 'continuous')
    expect(middle.sections.map((section) => [section.id, section.start])).toEqual([
      ['n9', undefined],
      ['s1', 'continuous'],
    ])
    expect(withoutSection(middle, 'n9').sections).toEqual([{ ...first, start: 'continuous' }])
  })

  it('nesta seção muda só ela; no documento todo, o papel de todas', () => {
    const draft = { ...DEFAULT_PAGE_SETUP, orientation: 'landscape' as const, pageNumberStart: 5 }
    const list = { page: DEFAULT_PAGE_SETUP, sections: [first, second] }
    const only = withPageSetup(list, 0, draft, 'section')
    expect(only.sections[0]).toMatchObject({ id: 's1', orientation: 'landscape', pageNumberStart: 5 })
    expect(only.page.orientation).toBe('portrait')

    const all = withPageSetup(list, 0, draft, 'document')
    expect([all.page, ...all.sections].map((section) => section.orientation)).toEqual([
      'landscape',
      'landscape',
      'landscape',
    ])
    // O reinício da numeração fica só na seção em que foi pedido.
    expect(all.page.pageNumberStart).toBeUndefined()
  })

  it('desvincular copia a faixa herdada com os endereços da seção dona; vincular a solta', () => {
    const band = {
      ...header,
      center: [{ kind: 'text' as const, text: 'Título', pid: 'rId5:0:0', bold: false, italic: false }],
    }
    const inherited = { ...DEFAULT_PAGE_SETUP, headerBand: band }
    expect(isLinkedToPrevious(second, 1, 'header')).toBe(true)
    const own = withBandsLinked(second, inherited, 's2', 'header', false)
    expect(own.headerBand?.center[0]?.pid).toBe('s2~rId5:0:0')
    expect(isLinkedToPrevious(own, 1, 'header')).toBe(false)
    expect(withBandsLinked(own, inherited, 's2', 'header', true).headerBand).toBeNull()
  })

  it('o texto decide que seções valem e em que ordem; a biblioteca só guarda', () => {
    // A ordem é a do corpo, e não a da biblioteca; a entrada que o texto não usa
    // não vale — é a marca que um desfazer tirou, e pode voltar.
    const extra = { ...second, id: 'n9' }
    const resolved = resolveSections(['s2', null, 's1'], null, DEFAULT_PAGE_SETUP, [first, second, extra])
    expect(resolved.sections.map((section) => section.id)).toEqual(['s2', 's1'])
    const doc = {
      type: 'doc',
      content: [
        { type: 'paragraph', attrs: { sectionBreak: 's1' } },
        {
          type: 'bulletList',
          content: [{ type: 'listItem', content: [{ type: 'paragraph', attrs: { sectionBreak: 's2' } }] }],
        },
      ],
    }
    expect(marksOfJson(doc)).toEqual(['s1', 's2'])
  })

  it('a quebra nova só acrescenta à biblioteca, e o desfazer volta tudo como era', () => {
    const library = [first]
    const before = resolveSections(['s1'], null, DEFAULT_PAGE_SETUP, library)
    // Na última seção: a de baixo é uma entrada nova, apontada pelo documento.
    const plan = planSectionBreak(before, library, 1, 'oddPage')
    expect(plan.rename).toBeNull()
    const grown = [...library, ...plan.additions]
    const after = resolveSections(['s1', plan.upperId], plan.bodyId, DEFAULT_PAGE_SETUP, grown)
    expect(after.sections.map((section) => section.id)).toEqual(['s1', plan.upperId])
    expect(after.page.start).toBe('oddPage')
    // Desfeito o texto (sem a marca nova e sem o atributo), a biblioteca maior
    // não muda nada: a última seção volta a começar como antes.
    const undone = resolveSections(['s1'], null, DEFAULT_PAGE_SETUP, grown)
    expect(undone).toEqual(before)

    // No meio: a marca que fecha a seção partida passa a apontar a entrada nova.
    const middle = planSectionBreak(before, library, 0, 'continuous')
    expect(middle.rename?.from).toBe('s1')
    expect(middle.bodyId).toBeNull()
  })

  it('excluir a primeira quebra dá à seção de baixo as faixas que ela herdava', () => {
    const resolved = resolveSections(['s1', 's2'], null, DEFAULT_PAGE_SETUP, [first, second])
    const plan = planSectionDelete(resolved, 0)!
    expect(plan.removeId).toBe('s1')
    expect(plan.next.sections.map((section) => [section.id, section.headerBand])).toEqual([['s2', header]])
    const stored = storeSections(plan.next, null, DEFAULT_PAGE_SETUP, [first, second])
    // A entrada da excluída fica: o desfazer pode trazer a marca de volta.
    expect(stored.library.map((section) => section.id)).toEqual(['s1', 's2'])
  })
})
