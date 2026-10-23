import { describe, expect, it } from 'vitest'
import { getSchema } from '@tiptap/core'
import { EditorState, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { buildEditorExtensions } from './editor-extensions.js'
import { paginationKey, paginationPlugin } from './extensions/pagination.js'
import { createGapWriter, layoutPages, topMarginFor, type PageGaps } from './usePagination.js'

type Measured = Parameters<typeof layoutPages>[0]
type MetricsOf = Parameters<typeof layoutPages>[2]

const CONTENT_PX = 900
/** 8 pt after each paragraph, the Normal style's. */
const SPACE_AFTER_PX = 10.67
const SECOND_POSITION = 50

const metricsOf: MetricsOf = () => ({
  columns: 1,
  columnStepPx: 600,
  columnWidthPx: 600,
  separator: false,
  leftPx: 96,
  rightPx: 96,
  widthPx: 794,
  heightPx: 1123,
  contentPx: CONTENT_PX,
  topPx: 96,
  bottomPx: 96,
})

const flows = [{ height: CONTENT_PX, newSheet: false, parity: null, restart: null, columns: 1 }]

function block(top: number, height: number): Measured['blocks'][number] {
  return {
    top,
    height,
    breakpoints: [],
    isPageBreak: false,
    breakAfter: false,
    keepWithNext: false,
    section: 0,
  }
}

/** A full paragraph and a second one that no longer fits, a bottom margin apart. */
function twoParagraphs(): Measured {
  const secondTop = 880 + SPACE_AFTER_PX
  return {
    blocks: [block(0, 880), block(secondTop, 19)],
    targets: [
      { at: 0, start: { blockIndex: 0 }, nodes: [{ position: 0, natural: 0, collapse: 0 }] },
      {
        at: secondTop,
        start: { blockIndex: 1 },
        nodes: [{ position: SECOND_POSITION, natural: SPACE_AFTER_PX, collapse: SPACE_AFTER_PX }],
      },
    ],
    measuredNotes: new Map(),
    endnotes: [],
  }
}

describe('screen pagination', () => {
  it('pushes the block opening the sheet by the gap alone, since the margin above adds up to it', () => {
    const { gaps, layout } = layoutPages(twoParagraphs(), flows, metricsOf)

    expect(layout.pages).toBe(2)
    const gap = gaps.gaps.get(SECOND_POSITION)!
    expect(gap).toBeGreaterThan(0)
    // The paragraph has no top margin of its own: counting the one above again would push it
    // 10.67 px further on every measurement, without end.
    expect(gaps.written.get(SECOND_POSITION)).toBeCloseTo(gap)
  })

  it('writes the distance minus what the block above already adds', () => {
    expect(topMarginFor(SPACE_AFTER_PX, SPACE_AFTER_PX)).toBe(0)
    expect(topMarginFor(-489.33, SPACE_AFTER_PX)).toBeCloseTo(-500)
    expect(topMarginFor(30, SPACE_AFTER_PX)).toBeCloseTo(19.33)
  })
})

describe('gap writer', () => {
  const schema = getSchema(buildEditorExtensions(() => {}))
  const SHEET_GAP = 220

  function viewOf(...paragraphs: string[]): EditorView {
    const doc = schema.nodeFromJSON({
      type: 'doc',
      content: paragraphs.map((text) => ({ type: 'paragraph', content: [{ type: 'text', text }] })),
    })
    const view = {
      state: EditorState.create({ doc, plugins: [paginationPlugin()] }),
      dispatch: (transaction: Transaction) => {
        view.state = view.state.apply(transaction)
      },
    }
    return view as unknown as EditorView
  }

  function gapsAt(position: number): PageGaps {
    const at = new Map([[position, SHEET_GAP]])
    return { gaps: at, written: at, lines: new Map(), headers: [], columns: new Map() }
  }

  const marginsOf = (view: EditorView) => paginationKey.getState(view.state)!.find()

  it('writes the margin again when editing dropped it, even if the measurement did not change', () => {
    const view = viewOf('above', 'opening')
    const opening = view.state.doc.firstChild!.nodeSize
    const write = createGapWriter()
    write(view, gapsAt(opening))
    expect(marginsOf(view)).toHaveLength(1)

    // Replacing the node drops its decoration, and the new one sits at the same position.
    const retyped = schema.nodes['paragraph']!.create(null, schema.text('opening'))
    view.dispatch(view.state.tr.replaceWith(opening, view.state.doc.content.size, retyped))
    expect(marginsOf(view)).toHaveLength(0)

    write(view, gapsAt(opening))
    expect(marginsOf(view)).toHaveLength(1)
  })

  it('does not dispatch again for the same measurement', () => {
    const view = viewOf('above', 'opening')
    const opening = view.state.doc.firstChild!.nodeSize
    const write = createGapWriter()
    write(view, gapsAt(opening))
    const drawn = view.state

    write(view, gapsAt(opening))

    expect(view.state).toBe(drawn)
  })
})
