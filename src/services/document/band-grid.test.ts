import { describe, expect, it } from 'vitest'
import { bandSchema } from '@shared/schemas.js'
import { gridBox, type Band } from './band.js'
import { DEFAULT_PAGE_SETUP } from './model.js'

const band = (placement: Partial<Pick<Band, 'gridWidthMm' | 'gridIndentMm'>>): Band => ({
  left: [],
  center: [],
  right: [],
  rule: false,
  floats: [],
  rows: [],
  ...placement,
})

describe('gridBox', () => {
  const page = { ...DEFAULT_PAGE_SETUP, margins: { top: 20, right: 20, bottom: 20, left: 30 } }

  it('a table narrower than the text keeps its width and starts at its indent from the margin', () => {
    expect(gridBox(band({ gridWidthMm: 70, gridIndentMm: 90 }), page)).toEqual({
      width: '70.00mm',
      marginLeft: '110.00mm',
    })
  })

  it('without a width, the grid takes the whole band', () => {
    expect(gridBox(band({}), page)).toBeNull()
  })

  it('a band saved before the placement still opens', () => {
    expect(bandSchema.parse({ rows: [] }).gridWidthMm).toBeUndefined()
    expect(bandSchema.safeParse({ gridWidthMm: 80, gridIndentMm: -2 }).success).toBe(true)
  })
})
