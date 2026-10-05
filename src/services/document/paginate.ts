/**
 * Takes blocks **already measured** by the editor and returns the break points, without touching
 * the DOM.
 *
 * Positions are in **flow coordinates**: the height the block would have on a continuous strip,
 * without the gaps between sheets. Inserting the gaps changes the `offsetTop` of everything after;
 * in flow coordinates the computation is a single pass, and the renderer adds the gaps afterwards.
 */

export interface MeasuredBlock {
  /** In flow coordinates. */
  readonly top: number
  readonly height: number
  /** Tops of lines/items where a page may restart; empty means atomic. */
  readonly breakpoints: readonly number[]
  /** The `pageBreak` node, the break the user asked for with Ctrl+Enter. */
  readonly isPageBreak: boolean
  /** A break Word stored inside the paragraph: the sheet ends after the block. */
  readonly breakAfter: boolean
  /** `w:keepNext`: cannot stay alone at the foot of the page. */
  readonly keepWithNext: boolean
  /**
   * `w:keepLines`: the paragraph's lines stay together. The break points are still measured, and
   * only apply if the paragraph alone is taller than the sheet: then it cannot be kept together,
   * and Word breaks it too.
   */
  readonly keepLines?: boolean
  /**
   * `w:widowControl`, on by default in Word: breaks after the first line and before the last are
   * dropped, and paragraphs of up to three lines move whole.
   */
  readonly widowControl?: boolean
  /**
   * Height of the table header rows (`w:tblHeader`), which repeat at the top of each sheet the
   * table continues on: the new sheet has that much less for the rest of the table.
   */
  readonly repeatHeight?: number
  /**
   * Breaks outside the widow and orphan rule: the foot of an anchored capture, where LibreOffice
   * lets the paragraph's empty line move down while the frame stays. Also in `breakpoints`.
   */
  readonly freeBreakpoints?: readonly number[]
  /**
   * The empty line of an anchored capture's paragraph, which LibreOffice lets into the bottom
   * margin.
   */
  readonly hangingBottom?: number
  /** Index into `SectionFlow[]`. Absent means the first, a single-section document. */
  readonly section?: number
  /** `w:br w:type="column"`: the column ends after this block. */
  readonly columnBreakAfter?: boolean
  /** Footnotes whose references are in this block, in text order. */
  readonly notes?: readonly MeasuredNote[]
}

/**
 * The height does not depend on pagination: the body is measured outside the flow, at column width.
 */
export interface MeasuredNote {
  readonly id: string
  /** The foot of the reference's line, in flow coordinates. */
  readonly at: number
  readonly height: number
  /** Relative to the body top; the first is 0. */
  readonly lines: readonly number[]
}

/** Lines `[fromLine, toLine)`. */
export interface NoteSlice {
  readonly id: string
  readonly fromLine: number
  readonly toLine: number
}

/** Screen pixels, bands excluded: the usable height from `contentHeightMm`. */
export interface SectionFlow {
  readonly height: number
  /**
   * Also a continuous section with another paper or orientation, which Word treats as next page.
   */
  readonly newSheet: boolean
  /** `w:type` evenPage/oddPage. */
  readonly parity: 'even' | 'odd' | null
  /** `w:pgNumType/@w:start`. */
  readonly restart: number | null
  /** `w:cols`. Absent means one. */
  readonly columns?: number
}

/**
 * The editor is still a single strip: a column is drawn by **lifting** each column's first block to
 * the region top (negative `lift`) and shifting its blocks sideways. After the region, the next
 * block moves down to the foot of the tallest column (positive `lift`).
 */
export interface ColumnPlacement {
  readonly column: number
  /** In pixels, added to the block's gap. */
  readonly lift: number
}

/** The separator line is drawn in it. */
export interface ColumnRegion {
  /** Index among sheets with text. */
  readonly sheet: number
  /** Relative to the top of the sheet's text column. */
  readonly top: number
  readonly height: number
  readonly section: number
  readonly columns: number
}

export interface SheetPlan {
  /** Its paper, margins and bands apply. */
  readonly section: number
  /**
   * A blank sheet Word inserts so an even or odd section lands on the right sheet. It counts in the
   * numbering and has the next section's paper.
   */
  readonly blank: boolean
  readonly number: number
  /** The "Different first page" sheet. */
  readonly first: boolean
}

/** Including blank sheets. */
export interface PagePlan {
  readonly breaks: number[]
  readonly sheets: SheetPlan[]
  /** By index; other blocks have no entry. */
  readonly placements: Map<number, ColumnPlacement>
  readonly regions: ColumnRegion[]
  /** Per content sheet (index among sheets with text). */
  readonly notes: NoteSlice[][]
  /** Per content sheet, separator included; 0 without notes. */
  readonly noteHeights: number[]
}

export interface NoteFlow {
  readonly separator: number
}

/** The line after the last is the note's foot. */
export function noteLineTop(note: Pick<MeasuredNote, 'height' | 'lines'>, line: number): number {
  return line >= note.lines.length ? note.height : (note.lines[line] ?? 0)
}

/** Lines `[from, to)`. */
export function noteSpan(note: Pick<MeasuredNote, 'height' | 'lines'>, from: number, to: number): number {
  return noteLineTop(note, to) - noteLineTop(note, from)
}

function lineCount(note: MeasuredNote): number {
  return Math.max(note.lines.length, 1)
}

/** Each value is where a new page starts; an empty list is a one-page document. */
export function paginate(blocks: readonly MeasuredBlock[], pageHeight: number): number[] {
  return paginateSections(blocks, [{ height: pageHeight, newSheet: false, parity: null, restart: null }])
    .breaks
}

/**
 * A sheet has the height of the section opening it. A section starting on a new sheet breaks before
 * its first block, unless the sheet is empty; an even or odd section first gets a blank sheet when
 * the number does not match, as in Word.
 */
export function paginateSections(
  blocks: readonly MeasuredBlock[],
  sections: readonly SectionFlow[],
  noteFlow: NoteFlow = { separator: 0 },
): PagePlan {
  return new Paginator(blocks, sections, noteFlow.separator).run()
}

interface PendingNote {
  readonly note: MeasuredNote
  readonly from: number
}

type ColumnFill = ReturnType<typeof fillColumns>

const sectionOf = (block: MeasuredBlock | undefined, fallback: number): number => block?.section ?? fallback

class Paginator {
  private readonly breaks: number[] = []
  private readonly sheets: SheetPlan[] = []
  private readonly placements = new Map<number, ColumnPlacement>()
  private readonly regions: ColumnRegion[] = []
  private readonly notes: NoteSlice[][] = []
  private readonly noteHeights: number[] = []

  /**
   * A block moved down to the foot of a column region: if the sheet ends right before it, the move
   * does not apply, since the break places it.
   */
  private pendingLift: number | null = null

  /**
   * A sheet carries the notes whose references it carries. Long notes follow Word: the reference
   * line and at least the note's first line share a sheet, and the rest continues at the top of the
   * next notes area (`carry`).
   */
  private readonly footnotes: readonly MeasuredNote[]
  private nextNote = 0
  private carry: PendingNote[] = []

  private current: number
  private pageHeight: number

  /**
   * `pageStart` is where the sheet counts height from; `floor`, the last break. They only differ
   * when the sheet opens with a repeated table header: the count starts above the break, by the
   * header height, but nothing may move back before the break.
   */
  private pageStart = 0
  private floor = 0

  constructor(
    private readonly blocks: readonly MeasuredBlock[],
    private readonly sections: readonly SectionFlow[],
    private readonly separator: number,
  ) {
    this.footnotes = blocks.flatMap((block) => block.notes ?? [])
    this.current = sectionOf(blocks[0], 0)
    this.pageHeight = this.flowOf(this.current).height
  }

  /**
   * No page cap: on each round the index advances or `floor` grows strictly, and there are finitely
   * many such positions. A cap would stop the loop and pile the rest of the document on the last
   * sheet; the `pageHeight` guard protects against an invalid height.
   */
  run(): PagePlan {
    this.open(this.current)
    if (this.blocks.length === 0) return this.plan()

    let index = 0
    while (index < this.blocks.length) index = this.step(index)
    this.closeNotes()
    return this.plan()
  }

  private plan(): PagePlan {
    const { breaks, sheets, placements, regions, notes, noteHeights } = this
    return { breaks, sheets, placements, regions, notes, noteHeights }
  }

  private flowOf(section: number): SectionFlow {
    return (
      this.sections[section] ??
      this.sections.at(-1) ?? { height: 0, newSheet: false, parity: null, restart: null }
    )
  }

  /** The same block again when the sheet turned before it. */
  private step(index: number): number {
    const block = this.blocks[index]!
    const section = this.enterSection(block)

    if (this.pageHeight <= 0) return index + 1

    // A section with columns: its blocks go whole into this sheet's columns; what does not fit
    // opens the next sheet.
    const columns = this.flowOf(section).columns ?? 1
    if (columns > 1 && !block.isPageBreak) return this.layoutColumns(index, section, columns)

    // A manual break applies even with the page half full, so it comes before any height math.
    if (block.isPageBreak) {
      const after = block.top + block.height
      if (after > this.floor) this.cutAndRestart(after, this.current)
      return index + 1
    }

    const bottom = block.top + block.height
    if (
      bottom - this.pageStart + this.noteNeed(bottom) <=
      this.pageHeight + Math.min(block.hangingBottom ?? 0, this.pageHeight / 2)
    ) {
      // A break the paragraph carries applies after it, and not when nothing follows, or the
      // document would end with a blank sheet.
      if (block.breakAfter && index + 1 < this.blocks.length) this.cutAndRestart(bottom, this.current)
      return index + 1
    }

    if (this.breakInside(block)) return index
    return this.breakBefore(index, block)
  }

  /**
   * A new section starting on a new sheet breaks before its first block. If the sheet is still
   * empty (the previous section ended in a page break), there is nothing to break: the sheet
   * becomes the new section's.
   */
  private enterSection(block: MeasuredBlock): number {
    const section = sectionOf(block, this.current)
    if (section === this.current) return section
    this.current = section
    if (!this.flowOf(section).newSheet) return section
    if (block.top > this.floor) {
      this.cutAndRestart(block.top, section)
    } else {
      this.retarget(section)
      this.pageHeight = this.flowOf(section).height
    }
    return section
  }

  private breakInside(block: MeasuredBlock): boolean {
    const breakpoint = usableBreakpoints(block, this.pageHeight)
      .filter((at) => at > this.floor && at - this.pageStart + this.noteNeed(at) <= this.pageHeight)
      .at(-1)
    if (breakpoint === undefined) return false
    this.cut(breakpoint, this.current)
    this.floor = breakpoint
    // A header taller than half a sheet does not repeat: repeating it would leave no room for the
    // row it introduces.
    const repeat = block.repeatHeight ?? 0
    this.pageStart = repeat > 0 && repeat < this.pageHeight / 2 ? breakpoint - repeat : breakpoint
    return true
  }

  /** No line, item or table row fits: the break goes **before** the block that would overflow. */
  private breakBefore(index: number, block: MeasuredBlock): number {
    const { at, opening } = this.keptTogetherStart(index, block)
    if (at <= this.floor) {
      // With no break available, an atomic block keeps the sheet to itself, and the layout enlarges
      // the paper to hold it.
      const bottom = block.top + block.height
      const used = bottom - this.pageStart
      this.pageStart = this.floor = bottom
      if (index + 1 < this.blocks.length) this.cut(bottom, this.current, used)
      return index + 1
    }
    this.cutAndRestart(at, opening)
    return index
  }

  /** A heading alone at the foot of the page moves down with what it introduces. */
  private keptTogetherStart(index: number, block: MeasuredBlock): { at: number; opening: number } {
    let at = block.top
    let opening = this.current
    for (let candidate = index; candidate > 0; candidate--) {
      const previous = this.blocks[candidate - 1]
      if (previous === undefined || !previous.keepWithNext) break
      if (previous.top <= this.floor) break
      // Does not cross a section break that opens a sheet: the upper section's heading does not
      // move to the lower section's sheet.
      const previousSection = sectionOf(previous, this.current)
      if (previousSection !== this.current && this.flowOf(this.current).newSheet) break
      at = previous.top
      opening = previousSection
    }
    return { at, opening }
  }

  /**
   * The last sheet closes with the remaining notes; a note that still did not fit continues on
   * notes-only sheets after the text.
   */
  private closeNotes(): void {
    const end = this.blocks.reduce((bottom, block) => Math.max(bottom, block.top + block.height), 0)
    let used = Math.max(end - this.pageStart, 0)
    for (;;) {
      this.settleNotes(Number.POSITIVE_INFINITY, used)
      if (this.carry.length === 0 || this.pageHeight <= 0) break
      this.breaks.push(end)
      this.open(this.current)
      this.pageStart = this.floor = end
      used = 0
    }
  }

  /**
   * Numbered from the previous sheet, or from the restart when it is the section's first. Parity
   * only applies to the first sheet of a section that asks for it, never to the document's first.
   */
  private open(section: number): void {
    const previous = this.sheets.at(-1)
    const first = previous === undefined || previous.section !== section
    const flow = this.flowOf(section)
    let number = first && flow.restart !== null ? flow.restart : (previous?.number ?? 0) + 1
    if (
      previous !== undefined &&
      first &&
      flow.parity !== null &&
      (number % 2 === 0) !== (flow.parity === 'even')
    ) {
      this.sheets.push({ section, blank: true, number, first: false })
      number += 1
    }
    this.sheets.push({ section, blank: false, number, first })
  }

  /**
   * The sheet that just opened, still empty, becomes the starting section's: redone with the new
   * section's numbering and parity.
   */
  private retarget(section: number): void {
    const last = this.sheets.at(-1)
    if (last === undefined || last.section === section) return
    this.sheets.pop()
    while (this.sheets.at(-1)?.blank === true) this.sheets.pop()
    this.open(section)
  }

  private cut(at: number, section: number, used = at - this.pageStart): void {
    if (this.pendingLift !== null && at <= this.blocks[this.pendingLift]!.top)
      this.placements.delete(this.pendingLift)
    this.pendingLift = null
    this.settleNotes(at, used)
    this.breaks.push(at)
    this.open(section)
    this.pageHeight = this.flowOf(section).height
  }

  private cutAndRestart(at: number, section: number): void {
    this.cut(at, section)
    this.pageStart = this.floor = at
  }

  /** The ones carried over from the previous sheet and those of references up to `at`. */
  private pendingNotes(at: number): PendingNote[] {
    const list = [...this.carry]
    for (
      let next = this.nextNote;
      next < this.footnotes.length && this.footnotes[next]!.at <= at + 0.5;
      next++
    ) {
      list.push({ note: this.footnotes[next]!, from: 0 })
    }
    return list
  }

  /**
   * New notes whole, except the last, which only needs its first line. Grows with `at`, so "the
   * last break that fits" still holds. The continuation comes before the text, as in Word, and asks
   * for the whole rest up to half a sheet.
   */
  private noteNeed(at: number): number {
    const { carry, footnotes, nextNote } = this
    if (carry.length === 0 && (nextNote >= footnotes.length || footnotes[nextNote]!.at > at + 0.5)) return 0
    const fresh = this.pendingNotes(at).slice(carry.length)
    let carried = 0
    for (const item of carry) carried += noteSpan(item.note, item.from, lineCount(item.note))
    const first = carry[0]
    let need =
      this.separator +
      (first === undefined
        ? 0
        : Math.max(Math.min(carried, this.pageHeight / 2), noteSpan(first.note, first.from, first.from + 1)))
    fresh.forEach((item, position) => {
      need +=
        position === fresh.length - 1
          ? noteSpan(item.note, item.from, item.from + 1)
          : noteSpan(item.note, item.from, lineCount(item.note))
    })
    return need
  }

  /**
   * Notes that fit go whole, the first that does not is cut between lines, and the rest continues
   * on the next sheet.
   */
  private settleNotes(at: number, used: number): void {
    const list = this.pendingNotes(at)
    while (this.nextNote < this.footnotes.length && this.footnotes[this.nextNote]!.at <= at + 0.5)
      this.nextNote += 1
    this.carry = []
    const placed: NoteSlice[] = []
    let room = this.pageHeight - used - this.separator
    let height = 0
    for (const item of list) {
      if (this.carry.length > 0) {
        this.carry.push(item)
        continue
      }
      const total = lineCount(item.note)
      let to = item.from
      while (to < total && noteSpan(item.note, item.from, to + 1) <= room + 0.5) to += 1
      // At least one line on a sheet that got none: that is what lets a note taller than the sheet
      // end, one sheet at a time.
      if (to === item.from && placed.length === 0) to += 1
      if (to > item.from) {
        const span = noteSpan(item.note, item.from, to)
        placed.push({ id: item.note.id, fromLine: item.from, toLine: to })
        room -= span
        height += span
      }
      if (to < total) this.carry.push({ note: item.note, from: to })
    }
    this.notes[this.breaks.length] = placed
    this.noteHeights[this.breaks.length] = placed.length > 0 ? height + this.separator : 0
  }

  /**
   * Returns the first block left out. Whole blocks only, approximating Word, which breaks between
   * lines.
   */
  private layoutColumns(start: number, section: number, count: number): number {
    const { blocks } = this
    let end = start
    while (end < blocks.length && sectionOf(blocks[end], section) === section) end += 1

    const first = blocks[start]!
    const offset = first.top - this.pageStart
    // The region's notes come out of the column height; their area sits below, at sheet width (a
    // declared limitation: Word puts them under each column).
    const last = blocks[end - 1]!
    const available = this.pageHeight - offset - this.noteNeed(last.top + last.height)
    // A region starting mid-sheet that cannot hold even its first block moves to the next sheet.
    if (offset > 0 && first.height > available) {
      this.cutAndRestart(first.top, section)
      return start
    }

    const fill = this.columnFill(start, end, available, count, section)
    const height = this.placeColumns(fill, offset)
    this.regions.push({ sheet: this.breaks.length, top: offset, height, section, columns: count })
    return this.leaveColumns(fill, end, section, offset + height)
  }

  /** Before a continuous section on the same sheet the columns are balanced, as in Word. */
  private columnFill(
    start: number,
    end: number,
    available: number,
    count: number,
    section: number,
  ): ColumnFill {
    const { blocks } = this
    const fill = fillColumns(blocks, start, end, available, count)
    const next = blocks[end]
    const balances =
      fill.stop === end &&
      !fill.forced &&
      next !== undefined &&
      !this.flowOf(sectionOf(next, section)).newSheet
    if (!balances) return fill

    let low = Math.max(...blocks.slice(start, end).map((block) => block.height), 1)
    let high = available
    for (let step = 0; step < 24 && high - low > 0.5; step++) {
      const middle = (low + high) / 2
      const trial = fillColumns(blocks, start, end, middle, count)
      if (trial.stop === end) high = middle
      else low = middle
    }
    return fillColumns(blocks, start, end, high, count)
  }

  /**
   * Each column's first block rises to the region top; the others follow, since the strip stays the
   * same inside the column. Returns the tallest column's height.
   */
  private placeColumns(fill: ColumnFill, offset: number): number {
    let height = 0
    for (const column of fill.columns) {
      const top = this.blocks[column.from]!
      const last = this.blocks[column.to - 1]!
      height = Math.max(height, last.top + last.height - top.top)
      const drawn = top.top - this.pageStart
      const lift = column.index === 0 ? 0 : offset - drawn
      for (let at = column.from; at < column.to; at++) {
        this.placements.set(at, { column: column.index, lift: at === column.from ? lift : 0 })
      }
      this.pageStart -= lift
    }
    return height
  }

  /** `regionBottom` is the foot of the column region, from the sheet top. */
  private leaveColumns(fill: ColumnFill, end: number, section: number, regionBottom: number): number {
    const stop = fill.stop
    const after = this.blocks[stop]
    if (after === undefined) return stop

    const lastPlaced = this.blocks[stop - 1]!
    if (fill.forced || stop < end) {
      // Sheet full, or a page or column break in the last column.
      const at = fill.forced ? lastPlaced.top + lastPlaced.height : after.top
      this.cutAndRestart(at, sectionOf(after, section))
      return stop
    }

    // The section ended on this sheet: the next block moves down to the foot of the tallest column,
    // keeping the natural space it already had above it.
    const gap = Math.max(after.top - (lastPlaced.top + lastPlaced.height), 0)
    const lift = regionBottom + gap - (after.top - this.pageStart)
    this.placements.set(stop, { column: 0, lift })
    this.pageStart -= lift
    this.pendingLift = stop
    return stop
  }
}

/** Whole block by whole block. */
function fillColumns(
  blocks: readonly MeasuredBlock[],
  start: number,
  end: number,
  height: number,
  count: number,
): { stop: number; forced: boolean; columns: { index: number; from: number; to: number }[] } {
  const columns: { index: number; from: number; to: number }[] = [{ index: 0, from: start, to: start }]
  const column = (): { index: number; from: number; to: number } => columns.at(-1)!
  const done = (stop: number, forced: boolean) => ({
    stop,
    forced,
    columns: columns.filter((item) => item.to > item.from),
  })
  for (let at = start; at < end; at++) {
    const block = blocks[at]!
    const opened = blocks[column().from]!
    if (at > column().from && block.top + block.height - opened.top > height) {
      if (column().index + 1 >= count) return done(at, false)
      columns.push({ index: column().index + 1, from: at, to: at })
    }
    column().to = at + 1
    if (block.breakAfter === true && at + 1 < blocks.length) return done(at + 1, true)
    if (block.columnBreakAfter === true && at + 1 < end) {
      if (column().index + 1 >= count) return done(at + 1, true)
      columns.push({ index: column().index + 1, from: at + 1, to: at + 1 })
    }
  }
  return done(end, false)
}

/** By the keep-together rules. */
function usableBreakpoints(block: MeasuredBlock, pageHeight: number): readonly number[] {
  if (block.keepLines === true && block.height <= pageHeight) return []
  if (block.widowControl !== true) return block.breakpoints
  const free = block.freeBreakpoints ?? []
  const lines = block.breakpoints.filter((at) => !free.includes(at))
  const guarded = [...lines.slice(1, -1), ...free].sort((left, right) => left - right)
  // Taller than the sheet and with no break that respects the rule: break anyway, since the
  // alternative is a sheet stretched past the paper.
  return guarded.length === 0 && block.height > pageHeight ? block.breakpoints : guarded
}
