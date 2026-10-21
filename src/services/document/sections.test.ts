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

describe('sections', () => {
  it('a band the section does not declare comes from the previous one; a declared empty one does not inherit', () => {
    const empty = { ...header, left: [], center: [], right: [] }
    const [a, b, c] = effectiveSections({ ...DEFAULT_PAGE_SETUP, headerBand: empty }, [first, second])
    expect(a!.headerBand).toBe(header)
    expect(b!.headerBand).toBe(header)
    expect(c!.headerBand).toBe(empty)
  })

  it('a block belongs to the section of the next mark; after the last, to the body section', () => {
    expect(blockSections([null, 's1', null, 's2', null], [first, second])).toEqual([0, 0, 1, 1, 2])
    // A mark with an unknown id (a paragraph pasted from another document) is not a mark.
    expect(blockSections(['x9', null], [first])).toEqual([1, 1])
  })

  it('a mark on a list item closes the section with the list', () => {
    const paragraph = { attrs: { sectionBreak: 's1' }, isTextblock: true }
    const list = {
      attrs: {},
      descendants: (visit: (node: typeof paragraph) => boolean | void) => void visit(paragraph),
    }
    expect(sectionBreakIn(list)).toBe('s1')
    expect(sectionBreakIn({ attrs: { sectionBreak: 's2' }, isTextblock: true })).toBe('s2')
  })

  it('a continuous one opens a sheet only when the paper changes; even and odd require parity', () => {
    expect(startsNewSheet({ ...DEFAULT_PAGE_SETUP, start: 'continuous' }, DEFAULT_PAGE_SETUP)).toBe(false)
    expect(startsNewSheet(second, DEFAULT_PAGE_SETUP)).toBe(true)
    expect(startsNewSheet(DEFAULT_PAGE_SETUP, DEFAULT_PAGE_SETUP)).toBe(true)
    expect(startsNewSheet(DEFAULT_PAGE_SETUP, undefined)).toBe(false)
    expect(parityOf({ ...DEFAULT_PAGE_SETUP, start: 'oddPage' })).toBe('odd')
    expect(parityOf({ ...DEFAULT_PAGE_SETUP, start: 'evenPage' })).toBe('even')
  })

  it('each sheet counts within its section, and the blank one is not a cover', () => {
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

  it('a new break copies the split section upward, and the lower one starts as requested', () => {
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

  it('in this section it changes only that one; in the whole document, the paper of all', () => {
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
    // The numbering restart stays only in the section where it was asked for.
    expect(all.page.pageNumberStart).toBeUndefined()
  })

  it("unlinking copies the inherited band with the owning section's addresses; linking releases it", () => {
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

  it('the text decides which sections count and in what order; the library only stores', () => {
    // Body order, not library order; an entry the text does not use does not count: it is the mark
    // an undo removed, and it may come back.
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

  it('a new break only adds to the library, and undo restores everything as it was', () => {
    const library = [first]
    const before = resolveSections(['s1'], null, DEFAULT_PAGE_SETUP, library)
    // In the last section: the lower one is a new entry, pointed to by the document.
    const plan = planSectionBreak(before, library, 1, 'oddPage')
    expect(plan.rename).toBeNull()
    const grown = [...library, ...plan.additions]
    const after = resolveSections(['s1', plan.upperId], plan.bodyId, DEFAULT_PAGE_SETUP, grown)
    expect(after.sections.map((section) => section.id)).toEqual(['s1', plan.upperId])
    expect(after.page.start).toBe('oddPage')
    // With the text undone (no new mark and no attribute), the larger library changes nothing: the
    // last section starts as before again.
    const undone = resolveSections(['s1'], null, DEFAULT_PAGE_SETUP, grown)
    expect(undone).toEqual(before)

    // In the middle: the mark closing the split section points to the new entry.
    const middle = planSectionBreak(before, library, 0, 'continuous')
    expect(middle.rename?.from).toBe('s1')
    expect(middle.bodyId).toBeNull()
  })

  it('deleting the first break gives the lower section the bands it inherited', () => {
    const resolved = resolveSections(['s1', 's2'], null, DEFAULT_PAGE_SETUP, [first, second])
    const plan = planSectionDelete(resolved, 0)!
    expect(plan.removeId).toBe('s1')
    expect(plan.next.sections.map((section) => [section.id, section.headerBand])).toEqual([['s2', header]])
    const stored = storeSections(plan.next, null, DEFAULT_PAGE_SETUP, [first, second])
    // The deleted section's entry stays: undo may bring the mark back.
    expect(stored.library.map((section) => section.id)).toEqual(['s1', 's2'])
  })
})
