import { describe, expect, it } from 'vitest'
import {
  CELL_BORDER_SIDES,
  CellBorderStyle,
  DEFAULT_TABLE_DRAFT,
  MAX_BORDER_PT,
  MAX_TABLE_COLUMNS,
  MAX_TABLE_ROWS,
  NO_CELL_BORDERS,
  cellBordersFromAttr,
  cellBordersToAttr,
  cellBordersToCss,
  isValidTableDraft,
  isValidTableSize,
  resolvedColumnWidths,
  tableDraftFrom,
  cellLookPatch,
  withBorderOnSides,
} from './table-format.js'

describe('cell border and shading', () => {
  it('a cell with nothing declared opens the dialog with no side checked', () => {
    // A new cell: both attributes null. The dialog shows no declared side, and "Reset" offers all
    // four.
    expect(tableDraftFrom({ borders: null, shading: null })).toEqual({
      ...DEFAULT_TABLE_DRAFT,
      sides: { top: false, right: false, bottom: false, left: false },
    })
  })

  it('the attribute text has a single spelling', () => {
    // A single spelling for the fingerprint: `0,50 pt` and `0.5 pt` would regenerate the table.
    const borders = withBorderOnSides(NO_CELL_BORDERS, ['top', 'bottom'], {
      style: CellBorderStyle.Double,
      widthPt: 1.5,
      color: '#ff0000',
    })

    expect(cellBordersToAttr(borders)).toBe('top:double,1.5,#ff0000;bottom:double,1.5,#ff0000')
  })

  it('with no declared side the attribute is null, not empty text', () => {
    // Absent and null say the same; an empty string would say "there is a border, with no side".
    expect(cellBordersToAttr(NO_CELL_BORDERS)).toBeNull()
  })

  it('the attribute round-trips', () => {
    const attr = 'top:single,0.5,#000000;left:dotted,2,#123456'
    expect(cellBordersToAttr(cellBordersFromAttr(attr))).toBe(attr)
  })

  it('corrupted text does not bring the cell down', () => {
    // An attribute from pasted HTML or another version: at worst, a cell without borders.
    expect(cellBordersFromAttr('lixo')).toEqual(NO_CELL_BORDERS)
    expect(cellBordersFromAttr(42)).toEqual(NO_CELL_BORDERS)
    expect(cellBordersFromAttr('top:enfeite,1,#000000')).toEqual(NO_CELL_BORDERS)
  })

  it('a color outside the file format goes back to black', () => {
    // `w:color` only accepts six digits; Word reports `rgb(…)` as damaged.
    expect(cellBordersFromAttr('top:single,1,rgb(0,0,0)').top?.color).toBe('#000000')
  })

  it('a deliberately erased border becomes zero width in CSS', () => {
    // `w:val="nil"` erases the table border, which would reappear under `border-collapse`.
    const borders = withBorderOnSides(NO_CELL_BORDERS, ['bottom'], {
      style: CellBorderStyle.None,
      widthPt: 0.5,
      color: '#000000',
    })

    expect(cellBordersToCss(borders)).toBe('border-bottom:0')
  })

  it("the CSS comes out in points, the file's unit", () => {
    const borders = withBorderOnSides(NO_CELL_BORDERS, ['right'], {
      style: CellBorderStyle.Dashed,
      widthPt: 2.25,
      color: '#aabbcc',
    })

    expect(cellBordersToCss(borders)).toBe('border-right:2.25pt dashed #aabbcc')
  })

  it('a side unchecked in the dialog goes back to null, not to an erased border', () => {
    // "No opinion" lets the table border apply, "erased" removes it: absence and `w:nil`.
    const cell = { borders: 'top:single,0.5,#000000;right:single,0.5,#000000', shading: null }
    const before = tableDraftFrom(cell)
    const after = { ...before, sides: { ...before.sides, right: false } }

    expect(cellLookPatch(cell, before, after)).toEqual({ borders: 'top:single,0.5,#000000', shading: null })
  })

  it('shading comes out in lowercase, as the reader writes it', () => {
    const cell = { borders: null, shading: null }
    const before = tableDraftFrom(cell)
    const after = { ...before, shaded: true, shadingColor: '#D9D9D9' }
    expect(cellLookPatch(cell, before, after).shading).toBe('#d9d9d9')
  })

  it('the dialog opens with the four sides the cell has', () => {
    const draft = tableDraftFrom({
      borders: 'top:double,3,#00ff00;right:double,3,#00ff00',
      shading: '#d9d9d9',
      columnWidthMm: 42,
      headerRow: true,
    })

    expect(draft.borderStyle).toBe(CellBorderStyle.Double)
    expect(draft.borderWidthPt).toBe(3)
    expect(draft.sides).toEqual({ top: true, right: true, bottom: false, left: false })
    expect(draft.shaded).toBe(true)
    expect(draft.columnWidthMm).toBe(42)
    expect(draft.headerRow).toBe(true)
  })

  it('four sides, no more, no less', () => {
    // The `w:tcBorders` sides.
    expect(CELL_BORDER_SIDES).toEqual(['top', 'right', 'bottom', 'left'])
  })
})

describe('column widths as the screen draws them', () => {
  it('a table without a declared width splits the text column into equal parts', () => {
    // Like `table-layout: fixed` with `width: 100%`, a freshly inserted table.
    expect(resolvedColumnWidths([null, null, null], 600)).toEqual([200, 200, 200])
  })

  it('what remains is divided among the columns without a measure', () => {
    expect(resolvedColumnWidths([300, null, null], 600)).toEqual([300, 150, 150])
  })

  it('a width declared on all passes whole', () => {
    // A table wider than the column is the decision of whoever dragged it, and Word draws it that
    // way.
    expect(resolvedColumnWidths([400, 500], 600)).toEqual([400, 500])
  })

  it('a column never gets a measure the writer discards', () => {
    // A zero or negative width would make the writer drop the grid.
    expect(resolvedColumnWidths([null, null], 1).every((width) => width >= 1)).toBe(true)
  })
})

describe('table form validation', () => {
  it("a thickness outside the file's range does not pass", () => {
    // `w:sz` measures eighths of a point and fits in a byte: up to 31 pt.
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, borderWidthPt: 0 })).toBe(false)
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, borderWidthPt: MAX_BORDER_PT + 1 })).toBe(false)
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, borderWidthPt: Number.NaN })).toBe(false)
  })

  it('an unmeasured column width is valid: it simply does not change', () => {
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, columnWidthMm: null })).toBe(true)
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, columnWidthMm: 1 })).toBe(false)
  })

  it('a color the file does not accept fails the form', () => {
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, borderColor: 'vermelho' })).toBe(false)
  })

  it('table size has a ceiling, and it is not purism', () => {
    // Each cell is a measured paragraph: a thousand by a thousand would freeze the editor.
    expect(isValidTableSize(3, 3)).toBe(true)
    expect(isValidTableSize(0, 3)).toBe(false)
    expect(isValidTableSize(2.5, 3)).toBe(false)
    expect(isValidTableSize(MAX_TABLE_ROWS + 1, 3)).toBe(false)
    expect(isValidTableSize(3, MAX_TABLE_COLUMNS + 1)).toBe(false)
  })
})

describe('the dialog applies only what the person changed', () => {
  // Word writes `w:val="nil"` all the time, and the draft uses the first border for all four sides:
  // changing the shading must not erase the bottom border.
  const cell = { borders: 'top:none,0.5,#000000;bottom:single,1.5,#ff0000', shading: null }
  const before = tableDraftFrom(cell)

  it('changing only the shading does not touch the borders', () => {
    const after = { ...before, shaded: true, shadingColor: '#ffff00' }
    expect(cellLookPatch(cell, before, after)).toEqual({
      borders: 'top:none,0.5,#000000;bottom:single,1.5,#ff0000',
      shading: '#ffff00',
    })
  })

  it("changing the color changes each side's color and keeps each one's style and thickness", () => {
    const sides = { borders: 'top:single,0.5,#000000;bottom:double,3,#000000', shading: '#d9d9d9' }
    const before = tableDraftFrom(sides)
    const after = { ...before, borderColor: '#0000FF' }

    expect(cellLookPatch(sides, before, after)).toEqual({
      borders: 'top:single,0.5,#0000ff;bottom:double,3,#0000ff',
      shading: '#d9d9d9',
    })
  })

  it('unchecking a side removes only it, and checking a new side adds it with what the dialog shows', () => {
    const after = {
      ...before,
      sides: { ...before.sides, bottom: false, left: true },
    }

    expect(cellLookPatch(cell, before, after).borders).toBe('top:none,0.5,#000000;left:none,0.5,#000000')
  })

  it('nothing touched is nothing applied', () => {
    expect(cellLookPatch(cell, before, before)).toEqual(cell)
  })

  it('in another cell of the selection, what changed applies and the rest of it stays', () => {
    // The draft comes from the cursor's cell; the selected neighbour has its own borders, and only
    // the changed field reaches it.
    const neighbour = { borders: 'right:dotted,2,#00ff00', shading: '#eeeeee' }
    const after = { ...before, borderWidthPt: 1 }

    expect(cellLookPatch(neighbour, before, after)).toEqual({
      borders: 'right:dotted,1,#00ff00',
      shading: '#eeeeee',
    })
  })
})
