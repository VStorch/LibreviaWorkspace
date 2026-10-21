import { describe, expect, it } from 'vitest'
import { CellFormat, createSheet, getCell, setCell, type Sheet } from './model.js'
import {
  applyBorders,
  applyStyle,
  cellsIn,
  clearContents,
  deleteColumns,
  deleteRows,
  describeRange,
  parseRange,
  insertColumns,
  insertRows,
  rangeContains,
  singleCell,
  toggleStyle,
  writeText,
  writeValue,
  writeValues,
  editableText,
} from './edit.js'

const range = { fromRow: 0, fromColumn: 0, toRow: 1, toColumn: 1 }

/** A1:B2 filled and a loose value in D4. */
function filled(): Sheet {
  let sheet = createSheet('S')
  sheet = setCell(sheet, 0, 0, { value: 'a' })
  sheet = setCell(sheet, 0, 1, { value: 'b' })
  sheet = setCell(sheet, 1, 0, { value: 'c' })
  sheet = setCell(sheet, 1, 1, { value: 'd' })
  sheet = setCell(sheet, 3, 3, { value: 'longe' })
  return sheet
}

describe('range', () => {
  it('describes a single cell and a range', () => {
    expect(describeRange(singleCell(2, 1))).toBe('B3')
    expect(describeRange(range)).toBe('A1:B2')
  })

  it('reads back what describeRange writes, in any order, and rejects the rest', () => {
    expect(parseRange('A1:B2')).toEqual(range)
    expect(parseRange('b2:a1')).toEqual(range)
    expect(parseRange('C3')).toEqual(singleCell(2, 2))
    for (const text of ['', 'A1:', 'A1:B2:C3', '1A']) expect(parseRange(text)).toBeNull()
  })

  it('normalizes a range selected backwards', () => {
    // Dragging bottom-up is as common as the opposite.
    expect(describeRange({ fromRow: 5, fromColumn: 3, toRow: 1, toColumn: 1 })).toBe('B2:D6')
  })

  it('walks every cell', () => {
    expect([...cellsIn(range)]).toHaveLength(4)
  })
})

describe('formatting', () => {
  it('merges with the existing style instead of replacing it', () => {
    // Making it bold must not erase the cell's existing background.
    const painted = applyStyle(filled(), singleCell(0, 0), { background: '#ffff00' })
    const bolded = applyStyle(painted, singleCell(0, 0), { bold: true })

    expect(getCell(bolded, 0, 0)?.style).toEqual({ background: '#ffff00', bold: true })
  })

  it('applies to every cell in the range', () => {
    const styled = applyStyle(filled(), range, { align: 'center' })

    for (const { row, column } of cellsIn(range)) {
      expect(getCell(styled, row, column)?.style?.align).toBe('center')
    }
    // Outside the range, nothing changes.
    expect(getCell(styled, 3, 3)?.style).toBeUndefined()
  })

  it('formats empty cells, so typing starts formatted', () => {
    const styled = applyStyle(createSheet('S'), singleCell(5, 5), { format: CellFormat.Currency })

    expect(getCell(styled, 5, 5)?.style?.format).toBe(CellFormat.Currency)
  })

  it('preserves the value when changing the style', () => {
    const styled = applyStyle(filled(), singleCell(0, 0), { bold: true })

    expect(getCell(styled, 0, 0)?.value).toBe('a')
  })
})

describe('toggling an attribute', () => {
  it('turns on when the selection is mixed', () => {
    // Half bold: the expectation is to turn everything on, not to flip each cell.
    const half = applyStyle(filled(), singleCell(0, 0), { bold: true })
    const toggled = toggleStyle(half, range, 'bold')

    for (const { row, column } of cellsIn(range)) {
      expect(getCell(toggled, row, column)?.style?.bold).toBe(true)
    }
  })

  it('turns off when everything is already on', () => {
    const all = applyStyle(filled(), range, { bold: true })
    const toggled = toggleStyle(all, range, 'bold')

    for (const { row, column } of cellsIn(range)) {
      expect(getCell(toggled, row, column)?.style?.bold).toBeUndefined()
    }
  })

  it('leaves no empty style behind', () => {
    // Toggling on and off must give the cell back as it was, or the file grows with empty styles on
    // every click.
    const sheet = createSheet('S')
    const on = toggleStyle(sheet, singleCell(0, 0), 'bold')
    const off = toggleStyle(on, singleCell(0, 0), 'bold')

    expect(Object.keys(off.cells)).toEqual([])
  })
})

describe('borders', () => {
  it('applies and removes', () => {
    const bordered = applyBorders(filled(), singleCell(0, 0), ['top', 'bottom'])
    expect(getCell(bordered, 0, 0)?.style?.borders).toEqual(['top', 'bottom'])

    const cleared = applyBorders(bordered, singleCell(0, 0), [])
    expect(getCell(cleared, 0, 0)?.style?.borders).toBeUndefined()
  })
})

describe('clearing content', () => {
  it('clears the value and preserves the formatting', () => {
    // That is what Delete does in a spreadsheet: clears the data, keeps the formatting.
    const styled = applyStyle(filled(), singleCell(0, 0), { background: '#eee' })
    const cleared = clearContents(styled, singleCell(0, 0))

    expect(getCell(cleared, 0, 0)?.value).toBeUndefined()
    expect(getCell(cleared, 0, 0)?.style?.background).toBe('#eee')
  })

  it('removes a cell left with nothing', () => {
    const cleared = clearContents(filled(), singleCell(0, 0))
    expect(cleared.cells['A1']).toBeUndefined()
  })
})

describe('rows and columns', () => {
  it('inserts a row, shifting what comes after', () => {
    const shifted = insertRows(filled(), 1)

    expect(getCell(shifted, 0, 0)?.value).toBe('a')
    expect(getCell(shifted, 2, 0)?.value).toBe('c')
    expect(getCell(shifted, 4, 3)?.value).toBe('longe')
  })

  it('deletes a row and pulls the rest up', () => {
    const shifted = deleteRows(filled(), 0)

    expect(getCell(shifted, 0, 0)?.value).toBe('c')
    expect(getCell(shifted, 2, 3)?.value).toBe('longe')
  })

  it('inserts and deletes a column', () => {
    const inserted = insertColumns(filled(), 0)
    expect(getCell(inserted, 0, 1)?.value).toBe('a')

    const deleted = deleteColumns(inserted, 0)
    expect(getCell(deleted, 0, 0)?.value).toBe('a')
  })

  it('shifts the widths along with the columns', () => {
    // Otherwise inserting a column would leave the width on the wrong column.
    const sheet = { ...filled(), columnWidths: { 0: 200, 3: 60 } }
    const shifted = insertColumns(sheet, 0)

    expect(shifted.columnWidths[1]).toBe(200)
    expect(shifted.columnWidths[4]).toBe(60)
  })

  it('deletes several rows at once', () => {
    const shifted = deleteRows(filled(), 0, 2)

    expect(shifted.cells['A1']).toBeUndefined()
    expect(getCell(shifted, 1, 3)?.value).toBe('longe')
  })

  it('grows and shrinks the sheet along with it', () => {
    // Otherwise inserting would push the last row out of the visible area: the data would vanish
    // from the screen and stay in the file.
    const sheet = filled()

    expect(insertRows(sheet, 0, 3).rowCount).toBe(sheet.rowCount + 3)
    expect(deleteRows(sheet, 0, 3).rowCount).toBe(sheet.rowCount - 3)
    expect(insertColumns(sheet, 0).columnCount).toBe(sheet.columnCount + 1)
    expect(deleteColumns(sheet, 0).columnCount).toBe(sheet.columnCount - 1)
  })

  it('does not delete past the end of the sheet', () => {
    const sheet = { ...createSheet('S'), rowCount: 10, columnCount: 4 }

    expect(deleteRows(sheet, 8, 500).rowCount).toBe(8)
    expect(deleteColumns(sheet, 3, 500).columnCount).toBe(3)
    // Past the end there is nothing to delete: the sheet comes back intact.
    expect(deleteRows(sheet, 10, 1)).toBe(sheet)
  })

  it('moves the freeze panes along when the operation falls inside them', () => {
    const sheet = { ...createSheet('S'), frozenRows: 2, frozenColumns: 2 }

    expect(insertRows(sheet, 0).frozenRows).toBe(3)
    expect(deleteRows(sheet, 0).frozenRows).toBe(1)
    expect(insertColumns(sheet, 0).frozenColumns).toBe(3)
    // Below the frozen band, the freeze does not move.
    expect(insertRows(sheet, 2).frozenRows).toBe(2)
  })
})

describe('rangeContains', () => {
  it("accepts the rectangle's edges and rejects what is outside", () => {
    expect(rangeContains(range, 0, 0)).toBe(true)
    expect(rangeContains(range, 1, 1)).toBe(true)
    expect(rangeContains(range, 2, 1)).toBe(false)
    expect(rangeContains(range, 1, 2)).toBe(false)
  })

  it('works for a range dragged from bottom to top', () => {
    expect(rangeContains({ fromRow: 3, fromColumn: 3, toRow: 1, toColumn: 1 }, 2, 2)).toBe(true)
  })
})

describe('writeText', () => {
  it('stores the formula and leaves the value to recalculation', () => {
    const sheet = writeText(createSheet('S'), 0, 0, '=SOMA(A2:A3)')
    expect(getCell(sheet, 0, 0)).toEqual({ formula: '=SOMA(A2:A3)' })
  })

  it('recognizes the format of what was typed', () => {
    const sheet = writeText(createSheet('S'), 0, 0, 'R$ 10,50')
    expect(getCell(sheet, 0, 0)?.value).toBe(10.5)
    expect(getCell(sheet, 0, 0)?.style?.format).toBe(CellFormat.Currency)
  })

  it('does not erase a style the user chose by hand', () => {
    const painted = setCell(createSheet('S'), 0, 0, { value: 1, style: { bold: true } })
    expect(writeText(painted, 0, 0, '2').cells['A1']).toEqual({ value: 2, style: { bold: true } })
  })
})

describe('writeValue', () => {
  it('writes the value without turning it into a formula and keeps the cell style', () => {
    const styled = applyStyle(createSheet('S'), singleCell(0, 0), { bold: true })
    const sheet = writeValue(styled, 0, 0, '=A2')
    expect(getCell(sheet, 0, 0)).toEqual({ value: '=A2', style: { bold: true } })
    expect(getCell(writeValue(sheet, 0, 0, false), 0, 0)?.value).toBe(false)
    expect(getCell(writeValue(sheet, 0, 0, null), 0, 0)).toEqual({ style: { bold: true } })
  })
})

describe('writeValues', () => {
  it('writes several cells in a single copy, with the same rules as writeValue', () => {
    const styled = applyStyle(createSheet('S'), singleCell(1, 0), { bold: true })
    const sheet = writeValues(styled, [
      { row: 0, column: 0, value: '=A2' },
      { row: 1, column: 0, value: 'R$ 10' },
      { row: 2, column: 0, value: null },
    ])
    expect(getCell(sheet, 0, 0)).toEqual({ value: '=A2' })
    expect(getCell(sheet, 1, 0)?.value).toBe(10)
    expect(getCell(sheet, 1, 0)?.style?.bold).toBe(true)
    expect(getCell(sheet, 2, 0)).toBeUndefined()
    expect(styled.cells['A1']).toBeUndefined()
  })
})

describe('text with an apostrophe', () => {
  it('the apostrophe stores as text what would become a formula or number, and is not part of the value', () => {
    expect(getCell(writeText(createSheet('S'), 0, 0, "'=A2"), 0, 0)).toEqual({ value: '=A2' })
    expect(getCell(writeText(createSheet('S'), 0, 0, "'007"), 0, 0)).toEqual({ value: '007' })
    expect(getCell(writeText(createSheet('S'), 0, 0, "'12"), 0, 0)).toEqual({ value: '12' })
  })

  it('editing the cell shows the apostrophe, and confirming without change returns the same cell', () => {
    const literal = writeValue(createSheet('S'), 0, 0, '=A2')
    expect(editableText(getCell(literal, 0, 0))).toBe("'=A2")
    const again = writeText(literal, 0, 0, editableText(getCell(literal, 0, 0)))
    expect(getCell(again, 0, 0)).toEqual({ value: '=A2' })

    expect(editableText(getCell(writeText(createSheet('S'), 0, 0, 'Caneta'), 0, 0))).toBe('Caneta')
    expect(editableText(getCell(writeText(createSheet('S'), 0, 0, '=1+1'), 0, 0))).toBe('=1+1')
    expect(editableText(getCell(writeText(createSheet('S'), 0, 0, '3'), 0, 0))).toBe('3')
    expect(editableText(undefined)).toBe('')
  })
})
