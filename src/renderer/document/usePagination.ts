import { useCallback, useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'
import {
  noteLineTop,
  noteSpan,
  paginateSections,
  type MeasuredBlock,
  type MeasuredNote,
  type NoteSlice,
  type PagePlan,
  type SectionFlow,
  type SheetPlan,
} from '@services/document/paginate.js'
import { NoteKind } from '@services/document/notes.js'
import { effectiveAttrs } from '@services/document/style-cascade.js'
import type { StyleSheet } from '@services/document/styles.js'
import {
  contentHeightMm,
  contentInsetsMm,
  pageDimensionsMm,
  type PageSetup,
  type SectionSetup,
} from '@services/document/model.js'
import { mmToPx } from '@services/units.js'
import { NO_BANDS, type BandHeights } from '@services/document/band.js'
import {
  blockSections,
  parityOf,
  sectionBreakIn,
  startsNewSheet,
  columnGeometry,
  type SectionBlock,
} from '@services/document/sections.js'
import { applyPageGaps, type RepeatedHeader } from './extensions/pagination.js'
import { LINE_GAP_CLASS, measureLines } from './line-boxes.js'
import { noteBodyOf, type NoteBody } from './extensions/note-view.js'

/** The space between one sheet and the next, like a stack of paper. */
export const SHEET_GUTTER_PX = 28

/** A 12 pt line with the rule in the middle: Word's separator paragraph. */
export const NOTE_SEPARATOR_PX = 16

export interface NoteAreaItem {
  /** The body on screen (`note-view.ts`). */
  readonly key: string
  /**
   * The reference's order among all of the document's (`noteRefsOf`); paper finds the node by it.
   */
  readonly index: number
  readonly fromLine: number
  readonly toLine: number
  readonly clipTopPx: number
  readonly heightPx: number
}

/**
 * A sheet's notes area: footnotes at the foot of the text column, endnotes right after the last
 * block (and on the sheets after it).
 */
export interface NoteArea {
  readonly sheet: number
  readonly kind: 'footnote' | 'endnote'
  /** Sheet pixels, separator included. */
  readonly topPx: number
  readonly leftPx: number
  readonly widthPx: number
  readonly separator: 'normal' | 'continuation' | null
  readonly items: readonly NoteAreaItem[]
}

/** A constant, so `sameGaps` compares by value. */
const EMPTY_GAPS = new Map<number, number>()

export interface PageLayout {
  /** How many sheets to draw, blank sheets of even and odd sections included. */
  readonly pages: number
  readonly stackHeightPx: number
  readonly sheetTops: readonly number[]
  readonly sheetHeights: readonly number[]
  /** Slicing the block list at these points gives paper the same sheets as the screen. */
  readonly pageStarts: readonly PageStart[]
  /**
   * An anchored object's position comes from its paragraph, which only has a position after
   * pagination.
   */
  readonly anchors: readonly BlockAnchor[]
  readonly sheets: readonly SheetPlan[]
  readonly sheetWidths: readonly number[]
  /** The widest sheet's; the others are centered. */
  readonly stackWidthPx: number
  /** `pageStarts` only counts sheets with content; blank ones sit between them. */
  readonly contentSheets: readonly number[]
  /** Paper repeats the same offset (`print-source.ts`). */
  readonly columnMoves: readonly ColumnMove[]
  readonly columnLines: readonly ColumnLine[]
  readonly noteAreas: readonly NoteArea[]
}

/** A block in a section with columns: its column side and how much it moved up or down. */
export interface ColumnMove {
  readonly blockIndex: number
  readonly dx: number
  /** How much narrower the column is than the sheet's text column. */
  readonly narrowerPx: number
  readonly lift: number
  /** Paper writes the natural margin plus the offset. */
  readonly natural: number
  /** The previous block's bottom margin; see `collapsed`. */
  readonly collapse: number
}

/**
 * Vertical margins collapse: positive with positive takes the larger, and a negative one **adds**
 * to the positive. So a distance smaller than the previous bottom margin is only reached by
 * subtracting it.
 */
export function collapsed(distance: number, previous: number): number {
  return distance >= previous ? distance : distance - previous
}

/** In drawn sheet pixels. */
export interface ColumnLine {
  readonly sheet: number
  readonly leftPx: number
  readonly topPx: number
  readonly heightPx: number
}

/** The drawn sheet holding content sheet `index`. */
export function drawnSheet(layout: PageLayout, index: number): number {
  return layout.contentSheets[index] ?? index
}

interface SectionMetrics {
  readonly columns: number
  readonly columnStepPx: number
  readonly columnWidthPx: number
  readonly separator: boolean
  readonly leftPx: number
  readonly rightPx: number
  readonly widthPx: number
  readonly heightPx: number
  readonly contentPx: number
  readonly topPx: number
  readonly bottomPx: number
}

export interface PageStart {
  readonly blockIndex: number
  /** Index of the row or item opening the sheet, when the break is internal. */
  readonly childIndex?: number
  /** Position inside the paragraph content, the same `Node.cut` takes. */
  readonly offset?: number
  /** The sheet opens with the table header rows repeated. */
  readonly repeatHeader?: boolean
}

/** The break falls inside the block: the block starts on the previous sheet. */
export function isInternalStart(start: PageStart): boolean {
  return start.childIndex !== undefined || start.offset !== undefined
}

interface CutTarget {
  readonly at: number
  readonly start: PageStart
  readonly nodes: readonly { position: number; natural: number; collapse?: number }[]
  /** Resolved only if the break is chosen. */
  readonly line?: { readonly resolve: () => number | null; readonly block: number }
  /** A break between rows of a table with a header: what repeats at the top of the sheet. */
  readonly header?: () => RepeatedHeader
}

export interface BlockAnchor {
  readonly pageIndex: number
  /** In pixels, top margin included. */
  readonly topPx: number
}

/** What the editor passes to pagination besides the document and the sections. */
export interface PaginationOptions {
  readonly bands?: readonly BandHeights[]
  /**
   * Reading mode only turns off the push in the DOM: the math continues, and printing from inside
   * it yields the same sheets, because flow coordinates do not depend on applied gaps.
   */
  readonly paginated?: boolean
  /**
   * "Keep with next" may come from the style, and the block only carries what the paragraph
   * declares.
   */
  readonly styles?: StyleSheet | null
}

const INITIAL_LAYOUT: PageLayout = {
  pages: 1,
  stackHeightPx: 0,
  sheetTops: [0],
  sheetHeights: [],
  pageStarts: [],
  anchors: [],
  sheets: [{ section: 0, blank: false, number: 1, first: true }],
  sheetWidths: [],
  stackWidthPx: 0,
  contentSheets: [0],
  columnMoves: [],
  columnLines: [],
  noteAreas: [],
}

/**
 * The measurement is converted to **flow coordinates** before deciding: `offsetTop` already
 * includes the applied gaps, and subtracting them gives back the continuous strip height. Applying
 * the result does not change the next measurement's input, and the loop settles.
 */
export function usePagination(
  editor: Editor | null,
  /** With inherited bands resolved (`effectiveSections`); the last is the body's. */
  sections: readonly PageSetup[],
  /** A block finds its section by id. */
  declared: readonly SectionSetup[],
  revision: number,
  { bands = [], paginated = true, styles = null }: PaginationOptions = {},
): PageLayout {
  const [layout, setLayout] = useState<PageLayout>(INITIAL_LAYOUT)
  const writeGaps = useGapWriter()
  const bandsKey = bands.map((band) => `${band.headerMm}:${band.footerMm}`).join('|')

  useEffect(() => {
    if (editor === null) return undefined

    const element = editor.view.dom as HTMLElement // the resize observer target
    const metrics = sectionMetricsOf(sections, bands)
    const metricsOf = (section: number): SectionMetrics => metrics[section] ?? metrics.at(-1)!
    const flows: SectionFlow[] = sections.map((setup, index) => ({
      height: metricsOf(index).contentPx,
      newSheet: startsNewSheet(setup, sections[index - 1]),
      parity: parityOf(setup),
      restart: setup.pageNumberStart ?? null,
      columns: metricsOf(index).columns,
    }))

    const measureHidden = (): void => {
      const measured = new DocumentMeasurer(editor, element, styles).measure(declared)
      const result = layoutPages(measured, flows, metricsOf)
      // In reading mode the applied map empties too, or the next measurement would subtract a push
      // that does not exist.
      writeGaps(editor.view, paginated ? result.gaps : NO_PAGE_GAPS)
      setLayout(result.layout)
    }

    const measure = (): void => {
      // Line spacers are removed during measurement: they change where lines break, and the sheet
      // would end up with empty lines at the foot. The style comes back in the same frame, before
      // the browser paints.
      const lineGapsInDom = Array.from(element.querySelectorAll<HTMLElement>(`.${LINE_GAP_CLASS}`))
      for (const gap of lineGapsInDom) gap.style.display = 'none'
      try {
        measureHidden()
      } finally {
        for (const gap of lineGapsInDom) gap.style.display = 'inline-block'
      }
    }

    // One measurement per frame: fast typing would trigger dozens per second.
    let scheduled = 0
    const schedule = (): void => {
      if (scheduled !== 0) return
      scheduled = requestAnimationFrame(() => {
        scheduled = 0
        measure()
      })
    }

    schedule()

    const observer = new ResizeObserver(schedule)
    observer.observe(element)
    return () => {
      observer.disconnect()
      if (scheduled !== 0) cancelAnimationFrame(scheduled)
    }
  }, [editor, sections, declared, revision, bandsKey, paginated, styles, writeGaps])

  return layout
}

type MetricsOf = (section: number) => SectionMetrics

/** The margin is a floor: a taller header pushes it down. */
function sectionMetricsOf(sections: readonly PageSetup[], bands: readonly BandHeights[]): SectionMetrics[] {
  return sections.map((setup, index) => {
    const heights = bands[index] ?? NO_BANDS
    const insets = contentInsetsMm(setup, heights)
    const { width, height } = pageDimensionsMm(setup)
    const columns = columnGeometry(setup)
    return {
      columns: columns.count,
      columnStepPx: mmToPx(columns.stepMm),
      columnWidthPx: mmToPx(columns.widthMm),
      separator: columns.separator,
      leftPx: mmToPx(setup.margins.left),
      rightPx: mmToPx(setup.margins.right),
      widthPx: mmToPx(width),
      heightPx: mmToPx(height),
      contentPx: mmToPx(contentHeightMm(setup, heights)),
      topPx: mmToPx(insets.top),
      bottomPx: mmToPx(insets.bottom),
    }
  })
}

/** What measuring writes to the DOM, and what the next one subtracts. */
interface PageGaps {
  readonly gaps: ReadonlyMap<number, number>
  /**
   * The gap **plus** the natural margin, which is what CSS reads; the gap is what the math
   * subtracts.
   */
  readonly written: ReadonlyMap<number, number>
  /** Gaps between lines of a split paragraph, by spacer position. */
  readonly lines: ReadonlyMap<number, number>
  readonly headers: readonly RepeatedHeader[]
  /** Lateral column offset, by block position. */
  readonly columns: ReadonlyMap<number, number>
}

const NO_PAGE_GAPS: PageGaps = {
  gaps: EMPTY_GAPS,
  written: EMPTY_GAPS,
  lines: EMPTY_GAPS,
  headers: [],
  columns: EMPTY_GAPS,
}

/**
 * In a `ref`, because the effect reruns on every key press and the next read must subtract what was
 * already pushed. The key is the node **position**: an internal break pushes a table row or list
 * item, which has no block index.
 */
function useGapWriter(): (view: EditorView, next: PageGaps) => void {
  const last = useRef<PageGaps>(NO_PAGE_GAPS)
  /** Compared by what they draw. */
  const lastHeaders = useRef('[]')

  return useCallback((view: EditorView, next: PageGaps) => {
    const headersKey = JSON.stringify(next.headers)
    if (
      sameGaps(last.current.columns, next.columns) &&
      sameGaps(last.current.gaps, next.gaps) &&
      sameGaps(last.current.written, next.written) &&
      sameGaps(last.current.lines, next.lines) &&
      lastHeaders.current === headersKey
    )
      return
    lastHeaders.current = headersKey
    last.current = next
    applyPageGaps(view, next.written, next.gaps, {
      lines: next.lines,
      headers: next.headers,
      columns: next.columns,
    })
  }, [])
}

type IndexedNote = MeasuredNote & { index: number }

interface Measurement {
  readonly blocks: readonly MeasuredBlock[]
  readonly targets: readonly CutTarget[]
  readonly measuredNotes: ReadonlyMap<string, IndexedNote>
  /** In document order; they enter the flow after the last block. */
  readonly endnotes: readonly IndexedNote[]
}

/** A document block and the element drawing it. */
interface DrawnBlock {
  readonly block: ProseMirrorNode
  readonly dom: HTMLElement
  readonly offset: number
  readonly blockIndex: number
  readonly top: number
}

type MeasuredLines = NonNullable<ReturnType<typeof measureLines>>

/**
 * By the **document**, not the DOM children, which have other indexes: `nodeDOM` links one to the
 * other. Footnotes go with their reference's block; endnotes, after the last block. `refIndex` is
 * the `noteRefsOf` order.
 */
class DocumentMeasurer {
  private accumulated = 0
  private refIndex = 0
  private readonly origin: number
  private readonly blocks: MeasuredBlock[] = []
  private readonly targets: CutTarget[] = []
  private readonly measuredNotes = new Map<string, IndexedNote>()
  private readonly endnotes: IndexedNote[] = []

  constructor(
    private readonly editor: Editor,
    element: HTMLElement,
    private readonly styles: StyleSheet | null,
  ) {
    this.origin = offsetTopOf(element)
  }

  measure(declared: readonly SectionSetup[]): Measurement {
    const { doc } = this.editor.state
    const marks: (string | null)[] = []
    doc.forEach((block) => marks.push(sectionBreakIn(block as unknown as SectionBlock)))
    const sectionOfBlock = blockSections(marks, declared)

    doc.forEach((block, offset, blockIndex) => {
      const section = sectionOfBlock[blockIndex] ?? 0
      const dom = this.editor.view.nodeDOM(offset)
      if (dom instanceof HTMLElement) this.measureBlock(block, dom, offset, blockIndex, section)
      else this.skipUndrawn(block, section)
    })

    const { blocks, targets, measuredNotes, endnotes } = this
    return { blocks, targets, measuredNotes, endnotes }
  }

  private skipUndrawn(block: ProseMirrorNode, section: number): void {
    block.descendants((child) => {
      if (child.type.name !== 'noteRef') return true
      this.refIndex += 1
      return false
    })
    this.blocks.push({
      top: 0,
      height: 0,
      breakpoints: [],
      isPageBreak: false,
      breakAfter: false,
      keepWithNext: false,
      section,
    })
  }

  private measureBlock(
    block: ProseMirrorNode,
    dom: HTMLElement,
    offset: number,
    blockIndex: number,
    section: number,
  ): void {
    this.accumulated += shiftOf(dom)
    const drawn: DrawnBlock = {
      block,
      dom,
      offset,
      blockIndex,
      top: offsetTopOf(dom) - this.origin - this.accumulated,
    }
    this.pushBlockTarget(drawn)

    // Only the outer table's rows: nested ones belong to cells.
    const table =
      dom instanceof HTMLTableElement ? dom : dom.querySelector<HTMLTableElement>(':scope > table')
    const children = cutChildrenOf(dom, table)
    const breakpoints: number[] = []
    // The foot of the line of a note reference inside it.
    const childTops: number[] = []

    // Paragraphs and headings break between lines. Lines measure from the block border, and its
    // flow top is already in `top`.
    const lines = block.isTextblock && children.length === 0 ? measureLines(this.editor.view, dom) : null
    if (lines !== null) this.pushLineTargets(lines, drawn, breakpoints)
    const capture = anchoredCaptureOf(drawn, lines)
    const header = tableHeaderOf(table)
    const internal =
      lines !== null
        ? lines.shift
        : this.pushChildTargets(children, table, header, drawn, { breakpoints, childTops })

    const effective = effectiveAttrs(block, this.styles)
    const height = dom.offsetHeight - internal
    const notes = this.measureNotes(drawn, height, { lineStarts: lines?.starts ?? null, children, childTops })

    this.blocks.push({
      top: drawn.top,
      height,
      breakpoints,
      isPageBreak: dom.hasAttribute('data-page-break'),
      breakAfter: dom.hasAttribute('data-break-after'),
      columnBreakAfter: dom.hasAttribute('data-column-break'),
      keepWithNext: effective['keepNext'] === true || /^H[1-6]$/.test(dom.tagName),
      keepLines: effective['keepLines'] === true,
      widowControl: lines !== null && effective['widowControl'] !== false,
      ...(header.repeatHeight > 0 ? { repeatHeight: header.repeatHeight } : {}),
      ...(capture.freeBreakpoints.length > 0 ? { freeBreakpoints: capture.freeBreakpoints } : {}),
      ...(capture.hangingBottom > 0 ? { hangingBottom: capture.hangingBottom } : {}),
      ...(notes.length > 0 ? { notes } : {}),
      section,
    })
    this.accumulated += internal
  }

  private pushBlockTarget({ dom, offset, blockIndex, top }: DrawnBlock): void {
    const before = this.blocks.at(-1)
    const previousDom = dom.previousElementSibling
    this.targets.push({
      at: top,
      start: { blockIndex },
      nodes: [
        {
          position: offset,
          natural: Math.max(top - (before === undefined ? 0 : before.top + before.height), 0),
          // A negative top margin adds up with it (`collapsed`).
          collapse:
            previousDom instanceof HTMLElement
              ? parseFloat(getComputedStyle(previousDom).marginBottom) || 0
              : 0,
        },
      ],
    })
  }

  private pushLineTargets(lines: MeasuredLines, drawn: DrawnBlock, breakpoints: number[]): void {
    lines.starts.forEach((start, index) => {
      const at = drawn.top + start
      breakpoints.push(at)
      this.targets.push({
        at,
        start: { blockIndex: drawn.blockIndex },
        nodes: [],
        line: { resolve: () => lines.positionOf(index), block: drawn.offset },
      })
    })
  }

  /** Returns how much the gaps applied inside the block stretched it. */
  private pushChildTargets(
    children: readonly HTMLElement[],
    table: HTMLTableElement | null,
    header: TableHeader,
    drawn: DrawnBlock,
    cuts: { readonly breakpoints: number[]; readonly childTops: number[] },
  ): number {
    let internal = 0
    children.forEach((child, childIndex) => {
      const cells = child instanceof HTMLTableRowElement ? Array.from(child.cells) : []
      const shift = cells.length > 0 ? shiftOf(cells[0]!) : shiftOf(child)
      // Padding grows the row downward; a margin already moved its top.
      const at =
        offsetTopOf(child) - this.origin - this.accumulated - internal - (cells.length === 0 ? shift : 0)
      cuts.childTops.push(at)
      if (childIndex > header.headerRows) {
        cuts.breakpoints.push(at)
        this.targets.push({
          at,
          start: { blockIndex: drawn.blockIndex, childIndex },
          nodes: (cells.length > 0 ? cells : [child]).map((target) => ({
            position: this.editor.view.posAtDOM(target, 0) - 1,
            natural:
              parseFloat(
                cells.length > 0 ? getComputedStyle(target).paddingTop : getComputedStyle(target).marginTop,
              ) - shiftOf(target),
          })),
          ...(table !== null && header.repeatHeight > 0
            ? {
                header: () =>
                  repeatedHeader(
                    this.editor,
                    table,
                    header.rows.slice(0, header.headerRows),
                    cells[0]!,
                    header.repeatHeight,
                  ),
              }
            : {}),
        })
      }
      internal += shift
    })
    return internal
  }

  private measureNotes(drawn: DrawnBlock, height: number, rows: ReferenceRows): MeasuredNote[] {
    const notes: MeasuredNote[] = []
    drawn.block.descendants((child, pos) => {
      if (child.type.name !== 'noteRef') return true
      const index = this.refIndex++
      const reference = this.editor.view.nodeDOM(drawn.offset + 1 + pos)
      const body = noteBodyOf(reference)
      if (body === undefined || !(reference instanceof HTMLElement)) return false
      const at = referenceBottom(drawn.dom, reference, drawn.top, height, rows)
      const measured = { id: body.key, at, index, ...measureNote(body) }
      this.measuredNotes.set(body.key, measured)
      if (child.attrs['kind'] === NoteKind.Endnote) this.endnotes.push(measured)
      else notes.push(measured)
      return false
    })
    return notes
  }
}

/** Table rows, list items or table of contents entries: where the block can be broken. */
function cutChildrenOf(dom: HTMLElement, table: HTMLTableElement | null): HTMLElement[] {
  if (table !== null) return Array.from(table.rows)
  if (dom.tagName === 'UL' || dom.tagName === 'OL') {
    return Array.from(dom.children).filter(
      (child): child is HTMLElement => child instanceof HTMLElement && child.tagName === 'LI',
    )
  }
  // A table of contents breaks between entries, as a list between items.
  if (dom.hasAttribute('data-toc')) {
    return Array.from(dom.children).filter((child): child is HTMLElement => child instanceof HTMLElement)
  }
  return []
}

/**
 * An anchored capture: the empty line after the frame (a 1lh `::after`) moves to the next sheet
 * when it does not fit, as in LibreOffice.
 */
function anchoredCaptureOf(
  { block, dom, top }: DrawnBlock,
  lines: MeasuredLines | null,
): { freeBreakpoints: number[]; hangingBottom: number } {
  const freeBreakpoints: number[] = []
  if (
    lines === null ||
    dom.querySelector(':scope > .node-image[data-anchored], :scope > img[data-anchored]') === null
  )
    return { freeBreakpoints, hangingBottom: 0 }
  // With text, the break between frame and line is free from the widow rule.
  const first = lines.starts[0]
  if (
    first !== undefined &&
    block.firstChild?.type.name === 'image' &&
    block.firstChild.attrs['anchored'] === true
  ) {
    freeBreakpoints.push(top + first)
  }
  // Without text, the empty line may be left at the foot of the sheet.
  const after = parseFloat(getComputedStyle(dom, '::after').height)
  return { freeBreakpoints, hangingBottom: Number.isFinite(after) && after > 0 ? after : 0 }
}

interface TableHeader {
  readonly rows: readonly HTMLTableRowElement[]
  readonly headerRows: number
  readonly repeatHeight: number
}

/**
 * Header rows (`w:tblHeader`) at the start of the table: breaking inside them, or right after,
 * would leave the header alone at the foot.
 */
function tableHeaderOf(table: HTMLTableElement | null): TableHeader {
  const rows = table !== null ? Array.from(table.rows) : []
  let headerRows = 0
  while (
    headerRows < rows.length - 1 &&
    rows[headerRows]!.cells.length > 0 &&
    Array.from(rows[headerRows]!.cells).every((cell) => cell.tagName === 'TH')
  ) {
    headerRows += 1
  }
  const lastHeader = rows[headerRows - 1]
  const repeatHeight =
    lastHeader === undefined ? 0 : offsetTopOf(lastHeader) + lastHeader.offsetHeight - offsetTopOf(rows[0]!)
  return { rows, headerRows, repeatHeight }
}

function layoutPages(
  measured: Measurement,
  flows: readonly SectionFlow[],
  metricsOf: MetricsOf,
): { gaps: PageGaps; layout: PageLayout } {
  // Endnotes enter the flow after the last block, breaking between lines.
  const { blocks, endnotes } = measured
  const textBottom = blocks.reduce((bottom, block) => Math.max(bottom, block.top + block.height), 0)
  const lastSection = blocks.at(-1)?.section ?? 0
  const endnoteBlocks: MeasuredBlock[] = []
  let endnoteTop = textBottom + NOTE_SEPARATOR_PX
  for (const note of endnotes) {
    endnoteBlocks.push({
      top: endnoteTop,
      height: note.height,
      breakpoints: note.lines.slice(1).map((line) => endnoteTop + line),
      isPageBreak: false,
      breakAfter: false,
      keepWithNext: false,
      section: lastSection,
    })
    endnoteTop += note.height
  }

  const plan = paginateSections([...blocks, ...endnoteBlocks], flows, { separator: NOTE_SEPARATOR_PX })
  const stack = new SheetStack(plan, measured, metricsOf)
  plan.breaks.forEach((at, cut) => stack.cutAt(at, cut))
  stack.closeLast(endnoteBlocks.length > 0 ? endnoteTop : textBottom)
  stack.placeEndnotes(endnoteBlocks)
  return stack.result()
}

/**
 * Gap = what was left of the sheet + both margins + the space between papers. The written margin
 * **replaces** the block's natural margin (the last declaration wins), so the written value is the
 * gap plus the natural margin; the flow math subtracts only the gap.
 */
class SheetStack {
  private readonly gaps = new Map<number, number>()
  private readonly written = new Map<number, number>()
  private readonly lineGaps = new Map<number, number>()
  private readonly headers: RepeatedHeader[] = []
  private previous = 0
  private readonly pageStarts: PageStart[] = []
  private readonly sheetHeights: number[] = []
  private readonly noteAreas: NoteArea[] = []
  private readonly sheetStarts: number[] = []
  private readonly contentSheets: number[]
  /** The list and the breaks grow together: each sheet costs one lookup. */
  private readonly internalAt = new Map<number, CutTarget>()
  private cursor = 0

  constructor(
    private readonly plan: PagePlan,
    private readonly measured: Measurement,
    private readonly metricsOf: MetricsOf,
  ) {
    this.contentSheets = plan.sheets.flatMap((sheet, index) => (sheet.blank ? [] : [index]))
    for (const target of measured.targets) {
      if (
        (target.start.childIndex !== undefined || target.line !== undefined) &&
        !this.internalAt.has(target.at)
      ) {
        this.internalAt.set(target.at, target)
      }
    }
  }

  private get blocks(): readonly MeasuredBlock[] {
    return this.measured.blocks
  }

  private sheetOf(content: number): SheetPlan {
    return this.plan.sheets[this.contentSheets[content] ?? 0]!
  }

  /** A sheet's drawn height is the strip's plus the column offsets. */
  private liftsBetween(from: number, to: number): number {
    let sum = 0
    for (const [index, placement] of this.plan.placements) {
      const top = this.blocks[index]?.top
      if (top !== undefined && top >= from && top < to) sum += placement.lift
    }
    return sum
  }

  private blockTargetFrom(at: number): CutTarget | undefined {
    const { targets } = this.measured
    while (
      this.cursor < targets.length &&
      (targets[this.cursor]!.at < at || targets[this.cursor]!.line !== undefined)
    )
      this.cursor++
    return targets[this.cursor]
  }

  cutAt(at: number, cut: number): void {
    const ending = this.metricsOf(this.sheetOf(cut).section)
    const opening = this.metricsOf(this.sheetOf(cut + 1).section)
    const blanks = this.plan.sheets.slice(
      (this.contentSheets[cut] ?? 0) + 1,
      this.contentSheets[cut + 1] ?? 0,
    )
    const internal = this.internalAt.get(at)
    const position = internal?.line?.resolve() ?? null
    // Line not found (DOM replaced mid-measurement): defer to the next block.
    const target =
      internal !== undefined && (internal.line === undefined || position !== null)
        ? internal
        : this.blockTargetFrom(at)
    // `hangingBottom` fits in the bottom margin.
    const span = at - this.previous + this.liftsBetween(this.previous, at)
    const hung = Math.min(Math.max(span - ending.contentPx, 0), hangingAt(this.blocks, at))
    const used = span - hung
    const notesHeight = this.plan.noteHeights[cut] ?? 0
    this.sheetStarts.push(this.previous)
    this.footnoteArea(cut, ending, used)
    this.sheetHeights.push(Math.max(ending.heightPx, used + notesHeight + ending.topPx + ending.bottomPx))
    let skipped = 0
    for (const blank of blanks) {
      const height = this.metricsOf(blank.section).heightPx
      this.sheetHeights.push(height)
      skipped += height + SHEET_GUTTER_PX
    }
    const shift =
      Math.max(ending.contentPx - used, notesHeight, 0) -
      hung +
      ending.bottomPx +
      SHEET_GUTTER_PX +
      skipped +
      opening.topPx
    this.previous = this.openSheet(at, target, position, shift, opening)
  }

  /** Returns where the new sheet counts its height from. */
  private openSheet(
    at: number,
    target: CutTarget | undefined,
    position: number | null,
    shift: number,
    opening: SectionMetrics,
  ): number {
    if (target?.line !== undefined && position !== null) {
      // Paper slices the paragraph at the same character.
      this.pageStarts.push({ ...target.start, offset: position - target.line.block - 1 })
      this.lineGaps.set(position, shift)
      return at
    }
    if (target === undefined) {
      // A final explicit break still opens an empty sheet.
      this.pageStarts.push({ blockIndex: this.blocks.length })
      return at
    }
    // The repeated header lives in the gap; the flow math subtracts both.
    const header = target.header?.()
    const extra = header !== undefined && header.height < opening.contentPx / 2 ? header.height : 0
    if (header !== undefined && extra > 0) this.headers.push(header)
    this.pageStarts.push(extra > 0 ? { ...target.start, repeatHeader: true } : target.start)
    for (const node of target.nodes) {
      this.gaps.set(node.position, shift + extra)
      this.written.set(node.position, shift + extra + node.natural)
    }
    return at - extra
  }

  closeLast(bottom: number): void {
    const content = this.plan.breaks.length
    const last = this.metricsOf(this.sheetOf(content).section)
    const lastSpan = Math.max(bottom - this.previous, 0) + this.liftsBetween(this.previous, bottom + 1)
    const lastHung = Math.min(Math.max(lastSpan - last.contentPx, 0), hangingAt(this.blocks, bottom))
    const lastNotes = this.plan.noteHeights[content] ?? 0
    this.sheetStarts.push(this.previous)
    this.footnoteArea(content, last, lastSpan - lastHung)
    this.sheetHeights.push(
      Math.max(last.heightPx, lastSpan - lastHung + lastNotes + last.topPx + last.bottomPx),
    )
  }

  /** At the foot of the text column, or right after the text if the sheet stretched. */
  private footnoteArea(content: number, metrics: SectionMetrics, used: number): void {
    const slices = this.plan.notes[content] ?? []
    const items = slices.flatMap((slice) => this.itemOf(slice))
    if (items.length === 0) return
    const height = this.plan.noteHeights[content] ?? 0
    this.noteAreas.push({
      sheet: this.contentSheets[content] ?? content,
      kind: 'footnote',
      topPx: metrics.topPx + Math.max(metrics.contentPx - height, used),
      leftPx: metrics.leftPx,
      widthPx: metrics.widthPx - metrics.leftPx - metrics.rightPx,
      separator: items[0]!.fromLine > 0 ? 'continuation' : 'normal',
      items,
    })
  }

  private itemOf(slice: NoteSlice): NoteAreaItem[] {
    const note = this.measured.measuredNotes.get(slice.id)
    if (note === undefined) return []
    return [
      {
        key: slice.id,
        index: note.index,
        fromLine: slice.fromLine,
        toLine: slice.toLine,
        clipTopPx: noteLineTop(note, slice.fromLine),
        heightPx: noteSpan(note, slice.fromLine, slice.toLine),
      },
    ]
  }

  placeEndnotes(endnoteBlocks: readonly MeasuredBlock[]): void {
    endnoteBlocks.forEach((block, position) => {
      const note = this.measured.endnotes[position]!
      this.sheetStarts.forEach((start, content) => this.placeEndnote(block, note, position, start, content))
    })
  }

  private placeEndnote(
    block: MeasuredBlock,
    note: IndexedNote,
    position: number,
    start: number,
    content: number,
  ): void {
    const end = this.plan.breaks[content] ?? Number.POSITIVE_INFINITY
    const visible = [block.top, ...block.breakpoints]
      .map((at, line) => ({ at, line }))
      .filter(({ at }) => at >= start - 0.5 && at < end - 0.5)
    if (visible.length === 0) return
    const fromLine = visible[0]!.line
    const toLine = visible.at(-1)!.line + 1
    const metrics = this.metricsOf(this.sheetOf(content).section)
    const sheet = this.contentSheets[content] ?? content
    const item: NoteAreaItem = {
      key: note.id,
      index: note.index,
      fromLine,
      toLine,
      clipTopPx: noteLineTop(note, fromLine),
      heightPx: noteSpan(note, fromLine, toLine),
    }
    const area = this.noteAreas.find((candidate) => candidate.sheet === sheet && candidate.kind === 'endnote')
    if (area !== undefined) {
      this.noteAreas[this.noteAreas.indexOf(area)] = { ...area, items: [...area.items, item] }
      return
    }
    // A sheet that only continues notes does not repeat the separator.
    const opens = position === 0 && fromLine === 0
    this.noteAreas.push({
      sheet,
      kind: 'endnote',
      topPx: metrics.topPx + (visible[0]!.at - start) - (opens ? NOTE_SEPARATOR_PX : 0),
      leftPx: metrics.leftPx,
      widthPx: metrics.widthPx - metrics.leftPx - metrics.rightPx,
      separator: opens ? 'normal' : null,
      items: [item],
    })
  }

  /** The columns' vertical offset goes into the gap; the lateral one is a translation. */
  private placeColumns(): { columnShifts: Map<number, number>; columnMoves: ColumnMove[] } {
    const columnShifts = new Map<number, number>()
    const blockTargets = this.wholeBlockTargets()
    const columnMoves: ColumnMove[] = []
    for (const [index, placement] of this.plan.placements) {
      const node = blockTargets.get(index)?.nodes[0]
      if (node === undefined) continue
      const metrics = this.metricsOf(this.blocks[index]?.section ?? 0)
      const dx = placement.column * metrics.columnStepPx
      if (dx !== 0) columnShifts.set(node.position, dx)
      if (placement.lift !== 0) this.liftBlock(node, placement.lift)
      columnMoves.push({
        blockIndex: index,
        dx,
        narrowerPx:
          metrics.columns > 1
            ? metrics.widthPx - metrics.leftPx - metrics.rightPx - metrics.columnWidthPx
            : 0,
        lift: placement.lift,
        natural: node.natural,
        collapse: node.collapse ?? 0,
      })
    }
    return { columnShifts, columnMoves }
  }

  /** Each block's first break that does not start in its middle. */
  private wholeBlockTargets(): Map<number, CutTarget> {
    const blockTargets = new Map<number, CutTarget>()
    for (const cut of this.measured.targets) {
      const whole = cut.line === undefined && cut.start.childIndex === undefined
      if (whole && !blockTargets.has(cut.start.blockIndex)) blockTargets.set(cut.start.blockIndex, cut)
    }
    return blockTargets
  }

  private liftBlock(node: CutTarget['nodes'][number], lift: number): void {
    this.gaps.set(node.position, (this.gaps.get(node.position) ?? 0) + lift)
    const distance = (this.written.get(node.position) ?? node.natural) + lift
    this.written.set(node.position, collapsed(distance, node.collapse ?? 0))
  }

  private columnLines(): ColumnLine[] {
    const lines: ColumnLine[] = []
    for (const region of this.plan.regions) {
      const metrics = this.metricsOf(region.section)
      if (!metrics.separator) continue
      const sheet = this.contentSheets[region.sheet] ?? region.sheet
      for (let column = 1; column < region.columns; column++) {
        lines.push({
          sheet,
          leftPx:
            metrics.leftPx +
            column * metrics.columnStepPx -
            (metrics.columnStepPx - metrics.columnWidthPx) / 2,
          topPx: metrics.topPx + region.top,
          heightPx: region.height,
        })
      }
    }
    return lines
  }

  result(): { gaps: PageGaps; layout: PageLayout } {
    // An even or odd section only asks for the blank sheet before starting.
    const sheets = this.plan.sheets.slice(0, this.sheetHeights.length)
    const sheetWidths = sheets.map((sheet) => this.metricsOf(sheet.section).widthPx)
    const sheetTops: number[] = []
    let stackHeightPx = 0
    for (const height of this.sheetHeights) {
      sheetTops.push(stackHeightPx)
      stackHeightPx += height + SHEET_GUTTER_PX
    }
    const { columnShifts, columnMoves } = this.placeColumns()
    const { contentSheets, plan } = this

    return {
      gaps: {
        gaps: this.gaps,
        written: this.written,
        lines: this.lineGaps,
        headers: this.headers,
        columns: columnShifts,
      },
      layout: {
        pages: this.sheetHeights.length,
        stackHeightPx: stackHeightPx - SHEET_GUTTER_PX,
        sheetTops,
        sheetHeights: this.sheetHeights,
        pageStarts: this.pageStarts,
        anchors: anchorsFor(
          this.blocks,
          plan.breaks,
          (content) => ({
            sheet: contentSheets[content] ?? content,
            marginTopPx: this.metricsOf(this.sheetOf(content).section).topPx,
          }),
          (index) => plan.placements.get(index)?.lift ?? 0,
        ),
        sheets,
        sheetWidths,
        stackWidthPx: Math.max(0, ...sheetWidths),
        contentSheets,
        columnMoves,
        columnLines: this.columnLines(),
        noteAreas: this.noteAreas,
      },
    }
  }
}

/** The capture's empty line and the gap to the block opening the next sheet. */
function hangingAt(blocks: readonly MeasuredBlock[], at: number): number {
  const ending = blocks.filter((block) => block.height > 0 && block.top + block.height <= at + 0.5).at(-1)
  if (ending === undefined || (ending.hangingBottom ?? 0) <= 0) return 0
  return ending.hangingBottom! + Math.max(at - (ending.top + ending.height), 0)
}

function sameGaps(left: ReadonlyMap<number, number>, right: ReadonlyMap<number, number>): boolean {
  if (left.size !== right.size) return false
  for (const [index, gap] of left) {
    // A pixel is not worth a transaction: the measure wobbles, and the document would shake.
    if (!right.has(index) || Math.abs(right.get(index)! - gap) > 0.5) return false
  }
  return true
}

/** Breaks are increasing boundaries: a single sweep. */
function anchorsFor(
  blocks: readonly MeasuredBlock[],
  breaks: readonly number[],
  sheetOf: (content: number) => { readonly sheet: number; readonly marginTopPx: number },
  liftOf: (index: number) => number = () => 0,
): BlockAnchor[] {
  const anchors: BlockAnchor[] = []
  let page = 0
  let drift = 0

  blocks.forEach((block, index) => {
    while (page < breaks.length && block.top >= breaks[page]!) {
      page += 1
      drift = 0
    }
    drift += liftOf(index)
    const start = page === 0 ? 0 : breaks[page - 1]!
    const { sheet, marginTopPx } = sheetOf(page)
    anchors.push({ pageIndex: sheet, topPx: marginTopPx + (block.top - start) + drift })
  })

  return anchors
}

/** Adds up the offsetParents' origins: a row measures from the table. */
function offsetTopOf(node: HTMLElement): number {
  let top = 0
  let current: HTMLElement | null = node
  while (current !== null) {
    top += current.offsetTop
    current = current.offsetParent instanceof HTMLElement ? current.offsetParent : null
  }
  return top
}

/** The decoration follows model edits; the measurement reads the already mapped push. */
function shiftOf(node: HTMLElement): number {
  return Number(node.dataset.pageShift ?? 0)
}

/**
 * A copy of the header rows over the gap of the row opening the sheet, pinned to its first cell by
 * negative margins.
 */
function repeatedHeader(
  editor: Editor,
  table: HTMLTableElement,
  headerRows: readonly HTMLTableRowElement[],
  cell: HTMLTableCellElement,
  height: number,
): RepeatedHeader {
  const tableBox = table.getBoundingClientRect()
  const cellBox = cell.getBoundingClientRect()
  const scale = table.offsetWidth > 0 && tableBox.width > 0 ? tableBox.width / table.offsetWidth : 1
  const style = getComputedStyle(cell)
  const left = (cellBox.left - tableBox.left) / scale + cell.clientLeft + parseFloat(style.paddingLeft)
  const natural = parseFloat(style.paddingTop) - shiftOf(cell)
  const colgroup = table.querySelector(':scope > colgroup')?.outerHTML ?? ''
  const html =
    `<table class="${table.className}" style="width:${table.offsetWidth}px;margin:0">${colgroup}` +
    `<tbody>${headerRows.map(cleanHeaderRow).join('')}</tbody></table>`
  return {
    position: editor.view.posAtDOM(cell, 0),
    html,
    height,
    offsetTop: height + Math.max(natural, 0),
    offsetLeft: left,
  }
}

/** No selection, search, handle or gaps: copied, they would show on every sheet. */
function cleanHeaderRow(row: HTMLTableRowElement): string {
  const copy = row.cloneNode(true) as HTMLTableRowElement
  for (const transient of copy.querySelectorAll(
    '.column-resize-handle, .page-line-gap, .page-repeated-header',
  )) {
    transient.remove()
  }
  for (const element of [copy, ...copy.querySelectorAll<HTMLElement>('*')]) {
    element.classList.remove('selectedCell', 'search-hit', 'search-hit--current', 'ProseMirror-selectednode')
    if (element.hasAttribute('data-page-start')) {
      element.style.removeProperty('padding-top')
      element.style.removeProperty('margin-top')
      element.removeAttribute('data-page-start')
      element.removeAttribute('data-page-shift')
    }
  }
  return copy.outerHTML
}

interface ReferenceRows {
  readonly lineStarts: readonly number[] | null
  readonly children: readonly HTMLElement[]
  readonly childTops: readonly number[]
}

/** How far the sheet must go to take the reference along. */
function referenceBottom(
  block: HTMLElement,
  reference: HTMLElement,
  top: number,
  height: number,
  { lineStarts, children, childTops }: ReferenceRows,
): number {
  if (lineStarts !== null) {
    const box = block.getBoundingClientRect()
    const scale = block.offsetHeight > 0 && box.height > 0 ? box.height / block.offsetHeight : 1
    // The superscript's foot, not its top: it rises above its line.
    const y = (reference.getBoundingClientRect().bottom - box.top) / scale - 1
    const next = lineStarts.find((start) => start > y)
    return top + (next ?? height)
  }
  const index = children.findIndex((child) => child.contains(reference))
  if (index >= 0 && childTops[index + 1] !== undefined) return childTops[index + 1]!
  return top + height
}

/** The note body's height and each line's top, at its current width. */
function measureNote(body: NoteBody): { height: number; lines: number[] } {
  const element = body.body
  const height = element.offsetHeight
  if (height <= 0) return { height: 0, lines: [0] }
  const box = element.getBoundingClientRect()
  const scale = box.height > 0 ? box.height / height : 1
  const lines: number[] = []
  for (const child of Array.from(element.children)) {
    if (!(child instanceof HTMLElement)) continue
    const top = (child.getBoundingClientRect().top - box.top) / scale
    const measured = /^(P|H[1-6])$/.test(child.tagName) ? measureLines(body.view, child) : null
    if (measured === null || measured.starts.length === 0) lines.push(top)
    else for (const start of measured.starts) lines.push(top + start)
  }
  const sorted = [...new Set(lines.map((line) => Math.round(line * 100) / 100))]
    .filter((line) => line >= 0 && line < height)
    .sort((left, right) => left - right)
  // The first line starts at the body top: the top margin goes with it.
  sorted[0] = 0
  return { height, lines: sorted }
}
