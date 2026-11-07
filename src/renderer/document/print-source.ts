import type { Editor } from '@tiptap/react'
import { DOMSerializer, Fragment, Node as ProseMirrorNode } from '@tiptap/pm/model'
import { type PageSetup } from '@services/document/model.js'
import { pxToMm } from '@services/units.js'
import { bandFloatsOf, floatsOf, type FloatingObject } from '@services/document/floating.js'
import type { PrintFloat, PrintNoteArea, PrintPage } from '@services/document/print-pages.js'
import { drawListsForPrint } from './extensions/list-numbering.js'
import { sheetSetups } from '@services/document/sections.js'
import { RevisionView } from '@shared/types.js'
import { DELETION, INSERTION } from './extensions/track-changes.js'
import { isHiddenBlock, isHiddenInline, revisionViewOf } from './extensions/revision-view.js'
import { drawsNoteNumber, noteLabelsOf, noteRefsOf, numberNotesForPrint } from './extensions/note-ref.js'
import { NOTE_NUMBER_CLASS } from './extensions/note-view.js'
import {
  NOTE_SEPARATOR_PX,
  drawnSheet,
  isInternalStart,
  topMarginFor,
  type PageLayout,
  type PageStart,
} from './usePagination.js'

/**
 * Serializes **straight from the nodes**, not `getHTML()`: Word stores page breaks inside
 * paragraphs, and a `<div>` inside `<p>` would make the HTML parser move it out, misaligning
 * indexes between screen and paper.
 */
export function splitIntoPages(
  editor: Editor,
  layout: PageLayout,
  sections: readonly PageSetup[],
): PrintPage[] {
  const context = printContextOf(editor, layout)
  const cuts: PageStart[] = [{ blockIndex: 0 }, ...layout.pageStarts, { blockIndex: context.blocks.length }]
  const pages: PrintPage[] = []
  const setups = sheetSetups(sections, layout.sheets)
  const setupOf = (drawn: number): SheetSetup => {
    const found = setups[drawn]
    return found === undefined
      ? { setup: sections.at(-1)!, inSection: drawn + 1 }
      : { setup: found.page, inSection: found.inSection }
  }

  for (let cut = 0; cut < cuts.length - 1; cut++) {
    // Blank sheets of an even or odd section carry only the band, as in Word.
    const drawn = drawnSheet(layout, cut)
    while (pages.length < drawn) {
      const blank = setupOf(pages.length)
      pages.push({
        number: pages.length + 1,
        html: '',
        floats: bandFloats(blank.setup, blank.inSection, editor),
        ...blank,
        blank: true,
      })
    }
    pages.push(contentPage(context, cuts[cut]!, cuts[cut + 1]!, pages.length, setupOf(pages.length)))
  }

  return pages
}

interface SheetSetup {
  readonly setup: PageSetup
  readonly inSection: number
}

interface PrintContext {
  readonly editor: Editor
  readonly layout: PageLayout
  readonly serializer: DOMSerializer
  /** With list numbering written in: the serializer does not see screen decorations. */
  readonly blocks: readonly ProseMirrorNode[]
  readonly view: RevisionView
  readonly offsets: readonly number[]
  /** On screen note numbers are decorations; here they are written with the screen's labels. */
  readonly screenLabels: readonly string[]
  readonly labelOf: ReadonlyMap<ProseMirrorNode, string>
}

function printContextOf(editor: Editor, layout: PageLayout): PrintContext {
  const offsets: number[] = []
  editor.state.doc.forEach((_node: ProseMirrorNode, offset: number) => {
    offsets.push(offset)
  })
  const screenLabels = noteLabelsOf(editor.state)
  const labelOf = new Map<ProseMirrorNode, string>()
  noteRefsOf(editor.state.doc).forEach(({ node }, index) => {
    const label = screenLabels[index]
    if (label !== undefined) labelOf.set(node, label)
  })
  return {
    editor,
    layout,
    serializer: DOMSerializer.fromSchema(editor.schema),
    blocks: drawListsForPrint(editor.state.doc),
    view: revisionViewOf(editor.state),
    offsets,
    screenLabels,
    labelOf,
  }
}

/** @param sheet the index among drawn sheets, blank ones included. */
function contentPage(
  context: PrintContext,
  start: PageStart,
  end: PageStart,
  sheet: number,
  setup: SheetSetup,
): PrintPage {
  const { editor, layout, blocks } = context
  return {
    number: sheet + 1,
    html: pageHtml(context, start, end),
    floats: [
      ...anchoredFloats(
        blocks,
        layout,
        start.blockIndex + (isInternalStart(start) ? 1 : 0),
        end.blockIndex + (isInternalStart(end) ? 1 : 0),
        editor,
      ),
      ...bandFloats(setup.setup, setup.inSection, editor),
    ],
    notes: notesForPrint(editor, layout, sheet, context.view, context.screenLabels),
    columnLines: layout.columnLines
      .filter((line) => line.sheet === sheet)
      .map((line) => ({
        leftMm: pxToMm(line.leftPx),
        topMm: pxToMm(line.topPx),
        heightMm: pxToMm(line.heightPx),
      })),
    ...setup,
  }
}

function pageHtml(context: PrintContext, start: PageStart, end: PageStart): string {
  const holder = document.createElement('div')
  // Slicing first, in screen indexes; the view mode after.
  const fragments = blocksForView(slicePageBlocks(context.blocks, start, end), context.view)
  holder.appendChild(context.serializer.serializeFragment(Fragment.fromArray(fragments)))
  // Comments do not go to paper, as in Word with markup off.
  for (const anchor of holder.querySelectorAll('[data-comment-start], [data-comment-end]')) anchor.remove()
  numberNotesForPrint(holder, printedNoteLabels(fragments, context.labelOf))
  // The same mark as the screen decoration: a capture with text does not get the 1lh line.
  for (const paragraph of holder.querySelectorAll('p')) {
    if (
      paragraph.querySelector(':scope > img[data-anchored]') !== null &&
      (paragraph.textContent ?? '').trim() !== ''
    ) {
      paragraph.setAttribute('data-anchor-text', '')
    }
  }
  placeColumns(holder, start, end, context.layout)
  markSplitParagraphs(
    holder,
    start,
    end,
    end.offset !== undefined && isJustified(context.editor, context.offsets[end.blockIndex]),
  )
  markSplitRows(holder, context.blocks, start, end)
  return holder.innerHTML
}

/** Paper has no decorations: the number is written at the start of the body. */
function notesForPrint(
  editor: Editor,
  layout: PageLayout,
  sheet: number,
  view: RevisionView,
  labels: readonly string[],
): PrintNoteArea[] {
  const areas = layout.noteAreas.filter((area) => area.sheet === sheet)
  if (areas.length === 0) return []
  const serializer = DOMSerializer.fromSchema(editor.schema)
  const refs = noteRefsOf(editor.state.doc)
  return areas.map((area) => ({
    topMm: pxToMm(area.topPx),
    leftMm: pxToMm(area.leftPx),
    widthMm: pxToMm(area.widthPx),
    separator: area.separator,
    separatorMm: pxToMm(NOTE_SEPARATOR_PX),
    items: area.items.flatMap((item) => {
      const reference = refs[item.index]
      if (reference === undefined) return []
      const children: ProseMirrorNode[] = []
      reference.node.forEach((child) => children.push(child))
      const holder = document.createElement('div')
      holder.appendChild(serializer.serializeFragment(Fragment.fromArray(blocksForView(children, view))))
      const first = holder.firstElementChild
      if (drawsNoteNumber(reference.node) && first !== null && /^(P|H[1-6])$/.test(first.tagName)) {
        const number = document.createElement('span')
        number.className = NOTE_NUMBER_CLASS
        number.textContent = labels[item.index] ?? ''
        first.prepend(number)
      }
      return [{ html: holder.innerHTML, clipTopMm: pxToMm(item.clipTopPx), heightMm: pxToMm(item.heightPx) }]
    }),
  }))
}

/** The same numbers the screen pagination produced. */
function placeColumns(holder: HTMLElement, start: PageStart, end: PageStart, layout: PageLayout): void {
  if (layout.columnMoves.length === 0) return
  const moves = new Map(layout.columnMoves.map((move) => [move.blockIndex, move]))
  const last = isInternalStart(end) ? end.blockIndex : end.blockIndex - 1
  const children = Array.from(holder.children)
  for (let index = start.blockIndex; index <= last; index++) {
    const element = children[index - start.blockIndex]
    const move = moves.get(index)
    if (!(element instanceof HTMLElement) || move === undefined) continue
    if (move.narrowerPx !== 0) element.style.marginRight = `${move.narrowerPx}px`
    if (move.dx !== 0) element.style.transform = `translateX(${move.dx}px)`
    // The sheet's first block does not rise: the new sheet already puts it at the top.
    if (move.lift !== 0 && index !== start.blockIndex) {
      element.style.marginTop = `${topMarginFor(move.natural + move.lift, move.collapse)}px`
    }
  }
}

/** Slices lines and items without duplicating content or restarting numbered lists. */
export function slicePageBlocks(
  blocks: readonly ProseMirrorNode[],
  start: PageStart,
  end: PageStart,
): ProseMirrorNode[] {
  const fragments: ProseMirrorNode[] = []
  for (let index = start.blockIndex; index <= end.blockIndex && index < blocks.length; index++) {
    const fragment = sliceBlock(blocks[index]!, index, start, end)
    if (fragment === null) break
    fragments.push(fragment)
  }
  return fragments
}

/** `null` when the break falls before the block's first content: the sheet is over. */
function sliceBlock(
  block: ProseMirrorNode,
  index: number,
  start: PageStart,
  end: PageStart,
): ProseMirrorNode | null {
  const textFrom = index === start.blockIndex ? start.offset : undefined
  const textTo = index === end.blockIndex ? end.offset : undefined
  if (block.isTextblock && (textFrom !== undefined || textTo !== undefined)) {
    return sliceTextblock(block, textFrom, textTo)
  }
  return sliceChildren(block, index, start, end)
}

function sliceChildren(
  block: ProseMirrorNode,
  index: number,
  start: PageStart,
  end: PageStart,
): ProseMirrorNode | null {
  const first = index === start.blockIndex
  const last = index === end.blockIndex
  const split: SplitRows = { from: first ? start.cells : undefined, to: last ? end.cells : undefined }
  const from = first ? (start.childIndex ?? 0) : 0
  const to = last ? (end.childIndex ?? 0) + (split.to === undefined ? 0 : 1) : block.childCount
  if (last && to === 0) return null
  if (from === 0 && to === block.childCount && split.from === undefined && split.to === undefined)
    return block
  return partialBlock(block, { from, to, repeatHeader: first && start.repeatHeader === true, split })
}

/** The cells where the first row of the sheet continues, and those where the last row stops. */
interface SplitRows {
  readonly from: readonly (number | null)[] | undefined
  readonly to: readonly (number | null)[] | undefined
}

/** A finished cell comes out empty, without the paragraph that would add a line. */
function rowPiece(row: ProseMirrorNode, { from, to }: SplitRows): ProseMirrorNode {
  const cells: ProseMirrorNode[] = []
  row.forEach((cell, _offset, index) => {
    const size = cell.content.size
    const begin = from === undefined ? 0 : (from[index] ?? size)
    const stop = to === undefined ? size : (to[index] ?? size)
    cells.push(begin >= stop ? cell.copy(Fragment.empty) : cell.cut(begin, stop))
  })
  return row.copy(Fragment.fromArray(cells))
}

/** Broken between lines, at the character where the screen placed the spacer. */
function sliceTextblock(
  block: ProseMirrorNode,
  textFrom: number | undefined,
  textTo: number | undefined,
): ProseMirrorNode | null {
  return textTo === 0 ? null : block.cut(textFrom ?? 0, textTo ?? block.content.size)
}

/** Children `[from, to)`, with the table header repeated; a numbered list continues its count. */
function partialBlock(
  block: ProseMirrorNode,
  { from, to, repeatHeader, split }: { from: number; to: number; repeatHeader: boolean; split: SplitRows },
): ProseMirrorNode {
  const children: ProseMirrorNode[] = []
  block.forEach((child, _offset, childIndex) => {
    const repeated = repeatHeader && isHeaderRow(block, childIndex)
    if (!repeated && (childIndex < from || childIndex >= to)) return
    const piece: SplitRows = {
      from: childIndex === from ? split.from : undefined,
      to: childIndex === to - 1 ? split.to : undefined,
    }
    children.push(piece.from === undefined && piece.to === undefined ? child : rowPiece(child, piece))
  })
  const attrs =
    block.type.name === 'orderedList'
      ? { ...block.attrs, start: Number(block.attrs.start ?? 1) + from }
      : block.attrs
  return block.type.create(attrs, Fragment.fromArray(children), block.marks)
}

/**
 * With all markup, the same blocks; with simple and no markup, the final text; with Original, the
 * opposite. A block the screen hides does not go.
 */
export function blocksForView(blocks: readonly ProseMirrorNode[], view: RevisionView): ProseMirrorNode[] {
  if (view === RevisionView.All) return [...blocks]
  return blocks.flatMap((block) => nodeForView(block, view) ?? [])
}

const originalNoteRefs = new WeakMap<ProseMirrorNode, ProseMirrorNode>()

/** What the view mode hid gets no label. */
function printedNoteLabels(
  fragments: readonly ProseMirrorNode[],
  labelOf: ReadonlyMap<ProseMirrorNode, string>,
): string[] {
  const labels: string[] = []
  for (const fragment of fragments) {
    fragment.descendants((node) => {
      if (node.type.name !== 'noteRef') return true
      labels.push(labelOf.get(originalNoteRefs.get(node) ?? node) ?? '')
      return false
    })
  }
  return labels
}

function nodeForView(node: ProseMirrorNode, view: RevisionView): ProseMirrorNode | null {
  if (node.isInline) {
    if (isHiddenInline(node, view)) return null
    const marks = node.marks.filter((mark) => mark.type.name !== INSERTION && mark.type.name !== DELETION)
    if (marks.length === node.marks.length) return node
    const shown = node.mark(marks)
    if (node.type.name === 'noteRef') originalNoteRefs.set(shown, node)
    return shown
  }
  if (isHiddenBlock(node, view)) return null

  const children: ProseMirrorNode[] = []
  node.forEach((child) => {
    const shown = nodeForView(child, view)
    if (shown !== null) children.push(shown)
  })
  const attrs = { ...node.attrs }
  if ('markRevision' in attrs) attrs['markRevision'] = null
  if ('rowRevision' in attrs) attrs['rowRevision'] = null
  if (children.length === 0 && node.childCount > 0 && !node.isTextblock) {
    // An empty cell stays, because the row needs it.
    return node.type.name === 'tableCell' || node.type.name === 'tableHeader'
      ? node.type.createAndFill(attrs, null, node.marks)
      : null
  }
  return node.type.create(attrs, Fragment.fromArray(children), node.marks)
}

/** A header row: at the start of the table and with only `tableHeader` cells. */
function isHeaderRow(table: ProseMirrorNode, rowIndex: number): boolean {
  for (let index = 0; index <= rowIndex; index++) {
    const row = table.maybeChild(index)
    if (row === null || row.childCount === 0) return false
    let header = true
    row.forEach((cell) => {
      if (cell.type.name !== 'tableHeader') header = false
    })
    if (!header) return false
  }
  return true
}

/**
 * On screen the two pieces are a single paragraph: the upper one loses its space after, the lower
 * one its space before and first-line indent. The last justified line above gets the screen's
 * spacer: `text-align-last` would also apply before each `<br>`.
 */
function markSplitParagraphs(
  holder: HTMLElement,
  start: PageStart,
  end: PageStart,
  justified: boolean,
): void {
  const first = holder.firstElementChild
  if (start.offset !== undefined && first instanceof HTMLElement) {
    first.style.marginTop = '0'
    first.style.paddingTop = '0'
    first.style.textIndent = '0'
    first.dataset.continued = 'from'
  }
  const last = holder.lastElementChild
  if (end.offset !== undefined && last instanceof HTMLElement) {
    last.style.marginBottom = '0'
    last.style.paddingBottom = '0'
    if (justified) {
      const filler = document.createElement('span')
      filler.setAttribute('aria-hidden', 'true')
      filler.style.cssText = 'display:inline-block;width:100%;height:0;vertical-align:top'
      last.appendChild(filler)
    }
    last.dataset.continued = 'to'
  }
}

/** The same as `markSplitParagraphs`, in each cell of a split row cut in the middle of a paragraph. */
function markSplitRows(
  holder: HTMLElement,
  blocks: readonly ProseMirrorNode[],
  start: PageStart,
  end: PageStart,
): void {
  const first = holder.firstElementChild
  if (start.cells !== undefined && first instanceof HTMLTableElement) {
    const row = blocks[start.blockIndex]?.maybeChild(start.childIndex ?? 0)
    forEachCutCell(first.rows[0], row, start.cells, (cell) => {
      const piece = cell.querySelector<HTMLElement>(TEXTBLOCK_SELECTOR)
      if (piece === null) return
      piece.style.marginTop = '0'
      piece.style.textIndent = '0'
    })
  }
  const last = holder.lastElementChild
  if (end.cells !== undefined && last instanceof HTMLTableElement) {
    const row = blocks[end.blockIndex]?.maybeChild(end.childIndex ?? 0)
    forEachCutCell(last.rows[last.rows.length - 1], row, end.cells, (cell) => {
      const piece = Array.from(cell.querySelectorAll<HTMLElement>(TEXTBLOCK_SELECTOR)).at(-1)
      if (piece !== undefined) piece.style.marginBottom = '0'
    })
  }
}

const TEXTBLOCK_SELECTOR = 'p, h1, h2, h3, h4, h5, h6'

function forEachCutCell(
  element: HTMLTableRowElement | undefined,
  row: ProseMirrorNode | null | undefined,
  cells: readonly (number | null)[],
  mark: (cell: HTMLTableCellElement) => void,
): void {
  if (element === undefined || row === null || row === undefined) return
  cells.forEach((offset, index) => {
    const cell = element.cells[index]
    if (offset === null || cell === undefined || index >= row.childCount) return
    const cut = row.child(index).resolve(offset)
    if (cut.parent.isTextblock && cut.parentOffset > 0) mark(cell)
  })
}

function isJustified(editor: Editor, offset: number | undefined): boolean {
  if (offset === undefined) return false
  const dom = editor.view.nodeDOM(offset)
  return dom instanceof HTMLElement && getComputedStyle(dom).textAlign === 'justify'
}

/** Serialized here, where the ProseMirror schema is known. */
function anchoredFloats(
  blocks: readonly ProseMirrorNode[],
  layout: PageLayout,
  start: number,
  end: number,
  editor: Editor,
): PrintFloat[] {
  const floats: PrintFloat[] = []

  for (let index = start; index < end; index++) {
    const anchor = layout.anchors[index]
    const block = blocks[index]
    if (anchor === undefined || block === undefined) continue

    for (const object of floatsOf(block.attrs)) {
      floats.push({ object, anchorTopMm: pxToMm(anchor.topPx), ...contentHtmlOf(object, editor) })
    }
  }

  return floats
}

function bandFloats(page: PageSetup, pageNumber: number, editor: Editor): PrintFloat[] {
  return bandFloatsOf(page, pageNumber).map((item) => ({
    ...item,
    ...contentHtmlOf(item.object, editor),
  }))
}

function contentHtmlOf(object: FloatingObject, editor: Editor): { contentHtml?: string } {
  if (object.kind !== 'text') return {}

  try {
    const nodes = (object.content ?? []).map((node) => ProseMirrorNode.fromJSON(editor.schema, node))
    const holder = document.createElement('div')
    holder.appendChild(DOMSerializer.fromSchema(editor.schema).serializeFragment(Fragment.fromArray(nodes)))
    return { contentHtml: holder.innerHTML }
  } catch {
    // A box the schema does not recognize comes out empty, without breaking the export.
    return { contentHtml: '' }
  }
}
