import { describe, expect, it } from 'vitest'
import { Language } from '@shared/i18n/index.js'
import {
  cellRef,
  columnIndex,
  columnName,
  createEmptyWorkbook,
  createSheet,
  getCell,
  parseRef,
  setCell,
} from './model.js'

describe('column names', () => {
  it.each([
    [0, 'A'],
    [25, 'Z'],
    [26, 'AA'],
    [27, 'AB'],
    [51, 'AZ'],
    [52, 'BA'],
    [701, 'ZZ'],
    [702, 'AAA'],
  ])('column %i is %s', (index, name) => {
    // Excel's base 26 has no zero: after Z comes AA, not BA. The classic mistake here, and it only
    // shows at column 26.
    expect(columnName(index)).toBe(name)
  })

  it('maps the name back to the index', () => {
    for (const index of [0, 25, 26, 51, 52, 701, 702]) {
      expect(columnIndex(columnName(index))).toBe(index)
    }
  })

  it('rejects a name that is not a column', () => {
    expect(columnIndex('A1')).toBe(-1)
    expect(columnIndex('')).toBe(-1)
  })
})

describe('A1 references', () => {
  it('builds the reference from the position', () => {
    expect(cellRef(0, 0)).toBe('A1')
    expect(cellRef(9, 27)).toBe('AB10')
  })

  it('reads the reference back', () => {
    expect(parseRef('AB10')).toEqual({ row: 9, column: 27 })
    expect(parseRef(' b2 ')).toEqual({ row: 1, column: 1 })
  })

  it.each(['', 'A', '1', 'A0', '1A', 'A-1'])('rejects %o', (ref) => {
    expect(parseRef(ref)).toBeNull()
  })
})

describe('cell writing', () => {
  it('stores and reads back', () => {
    const sheet = setCell(createSheet('S'), 2, 1, { value: 42 })

    expect(getCell(sheet, 2, 1)).toEqual({ value: 42 })
    expect(sheet.cells['B3']).toEqual({ value: 42 })
  })

  it('does not change the original sheet', () => {
    const original = createSheet('S')
    setCell(original, 0, 0, { value: 1 })

    expect(original.cells).toEqual({})
  })

  it('removes a cell left with nothing', () => {
    // Clearing an area must not leave thousands of empty entries in the file: it would grow with
    // every clear, without the user understanding why.
    const withValue = setCell(createSheet('S'), 0, 0, { value: 'x' })
    const cleared = setCell(withValue, 0, 0, { value: '' })

    expect(Object.keys(cleared.cells)).toEqual([])
  })

  it('keeps a cell that only has formatting', () => {
    // A yellow background without text is content: the user painted it on purpose.
    const painted = setCell(createSheet('S'), 0, 0, { style: { background: '#ffff00' } })

    expect(Object.keys(painted.cells)).toEqual(['A1'])
  })

  it('keeps a cell that only has a formula', () => {
    const withFormula = setCell(createSheet('S'), 0, 0, { formula: '=SOMA(B1:B9)' })

    expect(Object.keys(withFormula.cells)).toEqual(['A1'])
  })
})

describe('empty workbook', () => {
  it('names its sheet in the interface language, as Excel does', () => {
    expect(createEmptyWorkbook().sheets.map((sheet) => sheet.name)).toEqual(['Planilha1'])
    expect(createEmptyWorkbook(Language.English).sheets.map((sheet) => sheet.name)).toEqual(['Sheet1'])
  })
})
