import { describe, expect, it } from 'vitest'
import { cellMarginsCss } from './table-look.js'

describe('table cell margin', () => {
  it('twips become CSS pixels, in padding order', () => {
    expect(cellMarginsCss('0 108 60 108')).toBe('0px 7.2px 4px 7.2px')
  })

  it("anything other than four measures is dropped, and the cell gets the model's margin", () => {
    expect(cellMarginsCss(null)).toBeNull()
    expect(cellMarginsCss('0 108')).toBeNull()
    expect(cellMarginsCss('0 -5 0 108')).toBeNull()
  })
})
