import { describe, expect, it } from 'vitest'
import { writeText } from '@services/spreadsheet/edit.js'
import { createSheet, getCell } from '@services/spreadsheet/model.js'
import { applyFill, previewFormula, readFillAnswer, readFormulaAnswer } from './sheet-answers.js'

describe('readFormulaAnswer', () => {
  it('accepts the formula alone, inside code fences or in the middle of a sentence', () => {
    const expected = { kind: 'formula', formula: '=SOMA(B2:B3)' }
    expect(readFormulaAnswer('=SOMA(B2:B3)')).toEqual(expected)
    expect(readFormulaAnswer('```excel\n=SOMA(B2:B3)\n```')).toEqual(expected)
    expect(readFormulaAnswer('Use `=SOMA(B2:B3)` na célula.')).toEqual(expected)
    expect(readFormulaAnswer('Aqui está:\n=SOMA(B2:B3).')).toEqual(expected)
  })

  it('an "=" in the sentence before the formula does not hide it: the first valid formula snippet wins', () => {
    expect(readFormulaAnswer('Se A1 = 3, use =SOMA(B2:B3)')).toEqual({
      kind: 'formula',
      formula: '=SOMA(B2:B3)',
    })
    expect(readFormulaAnswer('Se A1 = 3, use =XLOOKUP(A1;B1:B9)').kind).toBe('invalid')
  })

  it('without "=" the answer has no formula', () => {
    expect(readFormulaAnswer('SOMA(B2:B3)')).toEqual({ kind: 'missing' })
    expect(readFormulaAnswer('')).toEqual({ kind: 'missing' })
  })

  it('a function Librevia does not know or a misspelling comes with the problem', () => {
    const unknown = readFormulaAnswer('=XLOOKUP(A1;B1:B9;C1:C9)')
    expect(unknown).toMatchObject({ kind: 'invalid', formula: '=XLOOKUP(A1;B1:B9;C1:C9)' })
    expect(unknown.kind === 'invalid' && unknown.message).toContain('XLOOKUP')
    expect(readFormulaAnswer('=SOMA(A1,B1)').kind).toBe('invalid')
  })
})

describe('readFillAnswer', () => {
  it('reads one value per row, also inside fences', () => {
    expect(readFillAnswer('["a", 2, true, null]', 4)).toEqual(['a', 2, true, null])
    expect(readFillAnswer('```json\n["a","b"]\n```', 2)).toEqual(['a', 'b'])
  })

  it('refuses wrong size, wrong types and text that is not JSON', () => {
    expect(readFillAnswer('["a"]', 2)).toBeNull()
    expect(readFillAnswer('[["a"], {"b": 1}]', 2)).toBeNull()
    expect(readFillAnswer('a, b', 2)).toBeNull()
    expect(readFillAnswer('[1, 2', 2)).toBeNull()
  })
})

describe('applyFill', () => {
  it('writes values, and text starting with "=" goes in as text, not as a formula', () => {
    const target = { fromRow: 1, toRow: 3, fromColumn: 2, toColumn: 2 }
    const sheet = applyFill(createSheet('S'), target, ['=1+1', 7, null])
    expect(getCell(sheet, 1, 2)).toEqual({ value: '=1+1' })
    expect(getCell(sheet, 2, 2)).toEqual({ value: 7 })
    expect(getCell(sheet, 3, 2)).toBeUndefined()
  })
})

describe('previewFormula', () => {
  it('computes on a copy, without changing the open workbook', () => {
    const sheet = writeText(writeText(createSheet('S'), 0, 0, '2'), 1, 0, '3')
    const workbook = { sheets: [sheet], activeSheet: 0 }
    expect(previewFormula(workbook, 0, { row: 2, column: 0 }, '=SOMA(A1:A2)')?.value).toBe(5)
    expect(getCell(workbook.sheets[0]!, 2, 0)).toBeUndefined()
  })
})
