import { describe, expect, it } from 'vitest'
import { layoutPages, topMarginFor } from './usePagination.js'

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
