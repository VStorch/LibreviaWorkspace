import { describe, expect, it } from 'vitest'
import { parseClipboardText, pasteText, rangeAsText } from './clipboard.js'
import { createSheet, getCell, setCell, type Sheet } from './model.js'

function filled(): Sheet {
  let sheet = createSheet('S')
  sheet = setCell(sheet, 0, 0, { value: 'a' })
  sheet = setCell(sheet, 0, 2, { formula: '=1+1', value: 2 })
  sheet = setCell(sheet, 1, 1, { value: 'linha\ncom "aspas"' })
  return sheet
}

describe('rangeAsText', () => {
  it('copies what the cell shows, with a tab between columns and one line per row', () => {
    expect(rangeAsText(filled(), { fromRow: 0, fromColumn: 0, toRow: 0, toColumn: 2 })).toBe('a\t\t2')
  })

  it('quotes a cell with a line break or quotes, as Excel does', () => {
    const range = { fromRow: 1, fromColumn: 0, toRow: 1, toColumn: 1 }
    expect(rangeAsText(filled(), range)).toBe('\t"linha\ncom ""aspas"""')
  })

  it('accepts a backwards range', () => {
    expect(rangeAsText(filled(), { fromRow: 0, fromColumn: 2, toRow: 0, toColumn: 0 })).toBe('a\t\t2')
  })
})

describe('parseClipboardText', () => {
  it('splits rows and columns and ignores the trailing line break', () => {
    expect(parseClipboardText('a\tb\nc\td\n')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
    expect(parseClipboardText('a\r\nb')).toEqual([['a'], ['b']])
  })

  it('keeps empty cells and reads the quotes back', () => {
    expect(parseClipboardText('\t"linha\ncom ""aspas"""\t')).toEqual([['', 'linha\ncom "aspas"', '']])
  })

  it('is the inverse of rangeAsText', () => {
    const range = { fromRow: 0, fromColumn: 0, toRow: 1, toColumn: 2 }
    const text = rangeAsText(filled(), range)
    expect(parseClipboardText(text)).toEqual([
      ['a', '', '2'],
      ['', 'linha\ncom "aspas"', ''],
    ])
  })
})

describe('pasteText', () => {
  it('writes from the cell as if typed, and discards what goes past the sheet', () => {
    const sheet = createSheet('S')
    const corner = sheet.columnCount - 1
    const pasted = pasteText(sheet, 0, corner - 1, [
      ['=1+2', 'x', 'fora'],
      ['3', 'y'],
    ])
    expect(getCell(pasted, 0, corner - 1)?.formula).toBe('=1+2')
    expect(getCell(pasted, 0, corner)?.value).toBe('x')
    expect(getCell(pasted, 1, corner - 1)?.value).toBe(3)
    expect(Object.keys(pasted.cells)).toHaveLength(4)
  })
})
