import { describe, expect, it } from 'vitest'
import {
  paginate,
  paginateSections,
  type MeasuredBlock,
  type MeasuredNote,
  type SectionFlow,
  noteLineTop,
  noteSpan,
} from './paginate.js'

/** Stacked fixed-height blocks, in order, as the editor measures them. */
function stack(
  heights: readonly number[],
  marks: { pageBreak?: number[]; breakAfter?: number[]; keepNext?: number[] } = {},
): MeasuredBlock[] {
  let top = 0
  return heights.map((height, index) => {
    const block: MeasuredBlock = {
      top,
      height,
      breakpoints: [],
      isPageBreak: marks.pageBreak?.includes(index) ?? false,
      breakAfter: marks.breakAfter?.includes(index) ?? false,
      keepWithNext: marks.keepNext?.includes(index) ?? false,
    }
    top += height
    return block
  })
}

describe('pagination', () => {
  it('a document that fits on one sheet does not break', () => {
    expect(paginate(stack([100, 100, 100]), 1000)).toEqual([])
  })

  it('breaks before the block that would overflow the page', () => {
    // The break falls at the top of the block that does not fit, never in its middle.
    expect(paginate(stack([400, 400, 400]), 1000)).toEqual([800])
  })

  it('a manually requested break holds with the page half full', () => {
    const blocks = stack([100, 10, 100], { pageBreak: [1] })
    expect(paginate(blocks, 1000)).toEqual([110])
  })

  it('a manual break at the start of the document creates no empty page', () => {
    // A break before any content has no page to close.
    const blocks = stack([10, 100], { pageBreak: [0] })
    expect(paginate(blocks, 1000)).toEqual([10])
  })

  it('a heading is not left alone at the foot of the page', () => {
    // An overflowing heading goes down with its paragraph, like the PDF's `break-after: avoid`.
    const blocks = stack([600, 200, 100, 300], { keepNext: [2] })
    expect(paginate(blocks, 1000)).toEqual([800])
  })

  it('a heading does not drag the whole page after it', () => {
    // The chain stops at the first block, already at the top: moving them all down would leave the
    // sheet blank.
    const blocks = stack([400, 400, 400], { keepNext: [0, 1, 2] })
    expect(paginate(blocks, 1000)).toEqual([400])
  })

  it('a block taller than the page gets the sheet to itself', () => {
    // An image taller than the sheet opens one, overflows, and what follows opens another.
    const blocks = stack([100, 3000, 100])
    expect(paginate(blocks, 1000)).toEqual([100, 3100])
  })

  it('a giant block at the end of the document creates no blank sheet', () => {
    // With nothing after it, no empty page at the end.
    const blocks = stack([100, 3000])
    expect(paginate(blocks, 1000)).toEqual([100])
  })

  it('the break a paragraph carries ends the sheet after it', () => {
    // `w:br w:type="page"` inside a `w:r` becomes a block property.
    const blocks = stack([100, 100, 100], { breakAfter: [1] })
    expect(paginate(blocks, 1000)).toEqual([200])
  })

  it('a break carried by the last block creates no blank sheet', () => {
    const blocks = stack([100, 100], { breakAfter: [1] })
    expect(paginate(blocks, 1000)).toEqual([])
  })

  it('an invalid page height breaks nothing', () => {
    // Margins larger than the paper give a negative height, and the loop must not run forever.
    expect(paginate(stack([100, 100]), 0)).toEqual([])
    expect(paginate(stack([100, 100]), -50)).toEqual([])
  })

  it('a long document produces one break per full page', () => {
    const blocks = stack(Array.from({ length: 20 }, () => 250))
    // 250 × 4 = 1000 per page; 20 blocks make five pages, four breaks.
    expect(paginate(blocks, 1000)).toEqual([1000, 2000, 3000, 4000])
  })

  it('a document with more than five hundred sheets does not pile the rest onto the last', () => {
    // No page cap: a short sheet in a regular document does not pile the rest onto the last one.
    const blocks = stack(Array.from({ length: 700 }, () => 100))
    const breaks = paginate(blocks, 100)

    expect(breaks).toHaveLength(699)
    expect(breaks.at(-1)).toBe(69_900)
  })

  it('a block taller than the sheet gets it to itself, even by the thousands', () => {
    // The loop also advances here, otherwise it would hang on every key press.
    const blocks = stack(Array.from({ length: 600 }, () => 300))
    expect(paginate(blocks, 100)).toHaveLength(599)
  })
})

describe('cuts inside blocks', () => {
  function splittable(
    height: number,
    breakpoints: number[],
    marks: Partial<MeasuredBlock> = {},
  ): MeasuredBlock {
    return { ...stack([height])[0]!, breakpoints, ...marks }
  }

  it('a table cuts at the last row that fits', () => {
    expect(paginate([splittable(1500, [300, 600, 900, 1200])], 1000)).toEqual([900])
  })

  it('the repeated header takes up the next sheet', () => {
    // Rows and header of 100: the second sheet repeats the header and fits one row less.
    const rows = Array.from({ length: 24 }, (_, index) => (index + 1) * 100)
    expect(paginate([splittable(2500, rows, { repeatHeight: 100 })], 1000)).toEqual([1000, 1900])
  })

  it('a header larger than half a sheet is not repeated', () => {
    const rows = Array.from({ length: 24 }, (_, index) => (index + 1) * 100)
    expect(paginate([splittable(2500, rows, { repeatHeight: 600 })], 1000)).toEqual([1000, 2000])
  })

  it('a three-page table has two internal cuts', () => {
    expect(paginate([splittable(2500, [500, 1000, 1500, 2000])], 1000)).toEqual([1000, 2000])
  })

  it('a list cuts at the top of the item', () => {
    expect(paginate([splittable(1200, [400, 800])], 1000)).toEqual([800])
  })

  it('a giant atomic block stays on a single sheet', () => {
    expect(paginate(stack([3000]), 1000)).toEqual([])
  })

  it('an explicit break precedes internal cuts', () => {
    expect(
      paginate(
        [splittable(1500, [500, 1000], { isPageBreak: true }), { ...stack([100])[0]!, top: 1500 }],
        1000,
      ),
    ).toEqual([1500])
  })

  it('a break after holds after the last piece of the table', () => {
    expect(
      paginate(
        [splittable(1500, [500, 1000], { breakAfter: true }), { ...stack([100])[0]!, top: 1500 }],
        1000,
      ),
    ).toEqual([1000, 1500])
  })

  it('a heading goes with the table when not even the first row fits', () => {
    const blocks = stack([800, 100, 1500], { keepNext: [1] })
    blocks[2] = { ...blocks[2]!, breakpoints: [1200, 1500, 1800, 2100] }
    expect(paginate(blocks, 1000)).toEqual([800, 1800])
  })

  it('keepWithNext does not move back before an internal cut', () => {
    const blocks = stack([100, 2500, 100], { keepNext: [0, 1] })
    blocks[1] = { ...blocks[1]!, breakpoints: [600, 1100, 1600, 2100] }
    expect(paginate(blocks, 1000)).toEqual([600, 1600, 2600])
  })

  describe('paragraph cut between lines', () => {
    // A ten-line paragraph with 50 px lines: the breaks are the tops of lines 2 to 10.
    const paragraph = (top: number, lines = 10, extra: Partial<MeasuredBlock> = {}): MeasuredBlock => ({
      top,
      height: lines * 50,
      breakpoints: Array.from({ length: lines - 1 }, (_, index) => top + (index + 1) * 50),
      isPageBreak: false,
      breakAfter: false,
      keepWithNext: false,
      ...extra,
    })

    it('the sheet ends at the last line that fits, not before the paragraph', () => {
      // Without the break, the paragraph would move down whole and leave a 300 gap.
      expect(paginate([...stack([700]), paragraph(700)], 1000)).toEqual([1000])
    })

    it('a line that does not fit entirely moves down', () => {
      expect(paginate([...stack([720]), paragraph(720)], 1000)).toEqual([970])
    })

    it('keeping lines together moves the whole paragraph down', () => {
      expect(paginate([...stack([700]), paragraph(700, 10, { keepLines: true })], 1000)).toEqual([700])
    })

    it('keeping lines together gives way when the paragraph is larger than the sheet', () => {
      expect(paginate([paragraph(0, 30, { keepLines: true })], 1000)).toEqual([1000])
    })

    it('the heading stays with the first lines of the paragraph it introduces', () => {
      const blocks = [...stack([700, 100], { keepNext: [1] }), paragraph(800)]
      expect(paginate(blocks, 1000)).toEqual([1000])
    })

    it('widows and orphans: the last line does not move down alone, it takes the penultimate', () => {
      // Nine of the ten lines would fit; the tenth would be a widow at the top of the sheet.
      expect(paginate([...stack([550]), paragraph(550, 10, { widowControl: true })], 1000)).toEqual([950])
    })

    it('widows and orphans: the first line is not left alone at the foot', () => {
      // Only one line would fit: the whole paragraph moves down.
      expect(paginate([...stack([930]), paragraph(930, 10, { widowControl: true })], 1000)).toEqual([930])
    })

    it('a paragraph of two or three lines moves whole', () => {
      expect(paginate([...stack([920]), paragraph(920, 2, { widowControl: true })], 1000)).toEqual([920])
      expect(paginate([...stack([880]), paragraph(880, 3, { widowControl: true })], 1000)).toEqual([880])
    })

    it('without the control, a lone line is accepted', () => {
      expect(paginate([...stack([930]), paragraph(930, 10, { widowControl: false })], 1000)).toEqual([980])
    })

    it('a paragraph larger than the sheet respects the rule at both cuts', () => {
      // 25 lines on a 20-line sheet: 20 and 5 break no rule, and cutting at 1000 works.
      expect(paginate([paragraph(0, 25, { widowControl: true })], 1000)).toEqual([1000])
      // 21 lines: cutting at 20 would leave the 21st a widow, so 19 stay.
      expect(paginate([paragraph(0, 21, { widowControl: true })], 1000)).toEqual([950])
    })

    it('the foot of the anchored capture cuts even with widow control', () => {
      // A 900 frame and its empty line (50), from 60: the line moves down and the frame stays, as
      // in LibreOffice.
      const capture: MeasuredBlock = {
        top: 60,
        height: 950,
        breakpoints: [960],
        freeBreakpoints: [960],
        isPageBreak: false,
        breakAfter: false,
        keepWithNext: false,
        widowControl: true,
      }
      expect(paginate([...stack([60]), capture], 1000)).toEqual([960])
    })

    it("the capture's empty line is left at the foot of the sheet instead of moving down", () => {
      // A 900 frame + a 50 line from 60: the empty line goes into the bottom margin, and the next
      // block opens the new sheet.
      const capture = { ...stack([950])[0]!, top: 60, hangingBottom: 50 }
      const after = { ...stack([100])[0]!, top: 1010 }
      expect(paginate([...stack([60]), capture, after], 1000)).toEqual([1010])
    })
  })
})

describe('pagination by section (M9)', () => {
  const flow = (height: number, options: Partial<SectionFlow> = {}): SectionFlow => ({
    height,
    newSheet: false,
    parity: null,
    restart: null,
    ...options,
  })
  const inSections = (heights: readonly number[], sections: readonly number[]): MeasuredBlock[] =>
    stack(heights).map((block, index) => ({ ...block, section: sections[index] ?? 0 }))

  it('a next-page section cuts before its first block', () => {
    const plan = paginateSections(inSections([100, 100, 100], [0, 1, 1]), [
      flow(1000),
      flow(1000, { newSheet: true }),
    ])
    expect(plan.breaks).toEqual([100])
    expect(plan.sheets.map((sheet) => [sheet.section, sheet.number, sheet.first])).toEqual([
      [0, 1, true],
      [1, 2, true],
    ])
  })

  it('a continuous one stays on the same sheet, and the next sheet has the height of the section that opens it', () => {
    // The first section has a 300 usable sheet; the continuous one, 500 (other margins). The second
    // sheet opens in section 1 and holds both 250 blocks.
    const plan = paginateSections(inSections([200, 50, 250, 250], [0, 1, 1, 1]), [flow(300), flow(500)])
    expect(plan.breaks).toEqual([250])
    expect(plan.sheets.map((sheet) => sheet.section)).toEqual([0, 1])
  })

  it('an odd section that would land on an even sheet gets a blank sheet before it', () => {
    const plan = paginateSections(inSections([100, 100, 100], [0, 1, 2]), [
      flow(1000),
      flow(1000, { newSheet: true, restart: 1 }),
      flow(1000, { newSheet: true, parity: 'odd' }),
    ])
    expect(plan.breaks).toEqual([100, 200])
    expect(plan.sheets.map((sheet) => [sheet.section, sheet.number, sheet.blank])).toEqual([
      [0, 1, false],
      [1, 1, false],
      [2, 2, true],
      [2, 3, false],
    ])
  })

  it('an even section that already lands on an even sheet gets no blank sheet', () => {
    const plan = paginateSections(inSections([100, 100], [0, 1]), [
      flow(1000),
      flow(1000, { newSheet: true, parity: 'even' }),
    ])
    expect(plan.sheets.map((sheet) => sheet.blank)).toEqual([false, false])
  })

  it('with the sheet still empty, the new section takes the sheet instead of opening another', () => {
    // A manual page break before the section mark leaves the new sheet empty: Word does not draw a
    // second blank sheet.
    const blocks = inSections([100, 10, 100], [0, 0, 1]).map((block, index) =>
      index === 1 ? { ...block, isPageBreak: true } : block,
    )
    const plan = paginateSections(blocks, [flow(1000), flow(1000, { newSheet: true, restart: 5 })])
    expect(plan.breaks).toEqual([110])
    expect(plan.sheets.map((sheet) => [sheet.section, sheet.number, sheet.first])).toEqual([
      [0, 1, true],
      [1, 5, true],
    ])
  })
})

describe('columns (M9)', () => {
  const flow = (height: number, options: Partial<SectionFlow> = {}): SectionFlow => ({
    height,
    newSheet: false,
    parity: null,
    restart: null,
    ...options,
  })
  const inSections = (heights: readonly number[], sections: readonly number[]): MeasuredBlock[] =>
    stack(heights).map((block, index) => ({ ...block, section: sections[index] ?? 0 }))
  const columnsOf = (plan: ReturnType<typeof paginateSections>, count: number): number[] =>
    Array.from({ length: count }, (_, index) => plan.placements.get(index)?.column ?? -1)

  it('fills the first column, moves the rest to the second and only then opens a sheet', () => {
    const plan = paginateSections(stack([400, 400, 400, 400, 400]), [flow(1000, { columns: 2 })])
    expect(columnsOf(plan, 5)).toEqual([0, 0, 1, 1, 0])
    // The first block of the second column lifts by what the first occupied.
    expect(plan.placements.get(2)?.lift).toBe(-800)
    expect(plan.breaks).toEqual([1600])
    expect(plan.regions.map((region) => [region.sheet, region.top, region.height])).toEqual([
      [0, 0, 800],
      [1, 0, 400],
    ])
  })

  it('before a continuous section the columns balance, and the following text drops to the foot', () => {
    const plan = paginateSections(inSections([100, 100, 100, 100, 50], [0, 0, 0, 0, 1]), [
      flow(1000, { columns: 2 }),
      flow(1000),
    ])
    expect(columnsOf(plan, 4)).toEqual([0, 0, 1, 1])
    expect(plan.regions[0]?.height).toBe(200)
    // The block of the section below was at 400 on the strip; drawn, it sits at 200.
    expect(plan.placements.get(4)?.lift).toBe(0)
    expect(plan.placements.get(2)?.lift).toBe(-200)
    expect(plan.breaks).toEqual([])
  })

  it('a column break moves the rest to the next column', () => {
    const blocks = stack([100, 100, 100]).map((block, index) =>
      index === 0 ? { ...block, columnBreakAfter: true } : block,
    )
    const plan = paginateSections(blocks, [flow(1000, { columns: 2 })])
    expect(columnsOf(plan, 3)).toEqual([0, 1, 1])
    expect(plan.placements.get(1)?.lift).toBe(-100)
  })
})

describe('footnotes (M11)', () => {
  const page = [{ height: 1000, newSheet: false, parity: null, restart: null }] as const
  const SEPARATOR = 20
  /** A note of `lines` 20 px lines, with its reference at the foot of `at`. */
  const note = (id: string, at: number, lines = 1): MeasuredNote => ({
    id,
    at,
    height: lines * 20,
    lines: Array.from({ length: lines }, (_, line) => line * 20),
  })
  /** A paragraph of 20 px lines, breakable between them. */
  const paragraph = (top: number, lines: number, notes: MeasuredNote[] = []): MeasuredBlock => ({
    top,
    height: lines * 20,
    breakpoints: Array.from({ length: lines - 1 }, (_, line) => top + (line + 1) * 20),
    isPageBreak: false,
    breakAfter: false,
    keepWithNext: false,
    notes,
  })

  it('reserves the separator and the note at the foot of the sheet', () => {
    // 50 lines fit (1000); with a one-line note (20 + 20 separator) on the first, only 48.
    const plan = paginateSections([paragraph(0, 60, [note('a', 20)])], page, { separator: SEPARATOR })
    expect(plan.breaks).toEqual([960])
    expect(plan.notes[0]).toEqual([{ id: 'a', fromLine: 0, toLine: 1 }])
    expect(plan.noteHeights[0]).toBe(40)
    expect(plan.notes[1]).toEqual([])
  })

  it('a note piece goes from the top of one line to the next, and the last to the foot', () => {
    const measured = { height: 70, lines: [0, 22, 40] }
    expect(noteLineTop(measured, 1)).toBe(22)
    expect(noteLineTop(measured, 3)).toBe(70)
    expect(noteSpan(measured, 1, 3)).toBe(48)
    expect(noteSpan(measured, 0, 1)).toBe(22)
  })

  it('without a note, nothing changes', () => {
    const blocks = [paragraph(0, 60)]
    expect(paginateSections(blocks, page, { separator: SEPARATOR }).breaks).toEqual(paginate(blocks, 1000))
  })

  it('a note that does not fit takes the reference line to the next sheet', () => {
    // The reference is on line 49 (foot at 980): the text would fit, the note's first line would
    // not, so the line moves down and the note goes with it.
    const plan = paginateSections([paragraph(0, 60, [note('a', 980)])], page, { separator: SEPARATOR })
    expect(plan.breaks).toEqual([960])
    expect(plan.notes[0]).toEqual([])
    expect(plan.notes[1]).toEqual([{ id: 'a', fromLine: 0, toLine: 1 }])
  })

  it('a long note keeps its first line on the reference sheet and continues on the next', () => {
    const blocks = [paragraph(0, 30, [note('a', 400, 40)]), paragraph(600, 25)]
    const plan = paginateSections(blocks, page, { separator: SEPARATOR })
    // The last note only needs its first line: the text runs until 960, and the note keeps what is
    // left (20 px, one line).
    expect(plan.breaks).toEqual([960])
    expect(plan.notes[0]).toEqual([{ id: 'a', fromLine: 0, toLine: 1 }])
    // On the next sheet the continuation comes before everything.
    expect(plan.notes[1]).toEqual([{ id: 'a', fromLine: 1, toLine: 40 }])
  })

  it('a long continuation fills the foot of the following sheets, without taking the whole sheet', () => {
    // A 120-line note on the first line with plenty of text after: the continuation asks for half a
    // sheet on each, not one line per sheet.
    const plan = paginateSections([paragraph(0, 1, [note('a', 20, 120)]), paragraph(20, 150)], page, {
      separator: SEPARATOR,
    })
    const slices = plan.notes.map((sheet) => sheet.map((slice) => [slice.fromLine, slice.toLine]))
    expect(slices[0]).toEqual([[0, 1]])
    expect(slices[1]).toEqual([[1, 26]])
    expect(slices[2]).toEqual([[26, 51]])
    // No sheet exceeds the height: text + separator + notes.
    plan.notes.forEach((_, sheet) => {
      const start = sheet === 0 ? 0 : plan.breaks[sheet - 1]!
      const end = plan.breaks[sheet] ?? 3020
      expect(end - start + (plan.noteHeights[sheet] ?? 0)).toBeLessThanOrEqual(1000)
    })
    expect(plan.notes.flat().at(-1)?.toLine).toBe(120)
  })

  it('a note larger than the sheet continues on notes-only sheets at the end', () => {
    const plan = paginateSections([paragraph(0, 2, [note('a', 20, 120)])], page, { separator: SEPARATOR })
    expect(plan.breaks.length).toBe(2)
    expect(plan.notes.map((slices) => slices.map((slice) => [slice.fromLine, slice.toLine]))).toEqual([
      [[0, 47]],
      [[47, 96]],
      [[96, 120]],
    ])
    expect(plan.sheets).toHaveLength(3)
  })

  it('notes from several references add up, and the last may be cut', () => {
    const blocks = [paragraph(0, 45, [note('a', 20, 3), note('b', 880, 5)])]
    const plan = paginateSections(blocks, page, { separator: SEPARATOR })
    // Up to 900 of text: 900 + 20 + 60 for a + the first line of b (20) = 1000 fits; the next line
    // does not. The previous one goes whole.
    expect(plan.breaks).toEqual([900])
    expect(plan.notes[0]).toEqual([
      { id: 'a', fromLine: 0, toLine: 3 },
      { id: 'b', fromLine: 0, toLine: 1 },
    ])
    expect(plan.notes[1]).toEqual([{ id: 'b', fromLine: 1, toLine: 5 }])
  })

  it('each section reserves the notes of its own sheet', () => {
    const blocks = [
      { ...paragraph(0, 10, [note('a', 20)]), section: 0 },
      { ...paragraph(200, 10, [note('b', 220)]), section: 1 },
    ]
    const plan = paginateSections(
      blocks,
      [
        { height: 1000, newSheet: false, parity: null, restart: null },
        { height: 1000, newSheet: true, parity: null, restart: null },
      ],
      { separator: SEPARATOR },
    )
    expect(plan.breaks).toEqual([200])
    expect(plan.notes).toEqual([[{ id: 'a', fromLine: 0, toLine: 1 }], [{ id: 'b', fromLine: 0, toLine: 1 }]])
  })

  it('endnotes, as blocks after the text, cut between their lines', () => {
    // Measuring (`usePagination`) puts them in the flow: here they are regular blocks.
    const blocks = [paragraph(0, 40), paragraph(820, 20)]
    const plan = paginateSections(blocks, page, { separator: SEPARATOR })
    expect(plan.breaks).toEqual([1000])
  })
})
