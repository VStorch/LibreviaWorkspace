import { describe, expect, it } from 'vitest'
import { cellRef, columnIndex, columnName, createSheet, getCell, parseRef, setCell } from './model.js'

describe('nomes de coluna', () => {
  it.each([
    [0, 'A'],
    [25, 'Z'],
    [26, 'AA'],
    [27, 'AB'],
    [51, 'AZ'],
    [52, 'BA'],
    [701, 'ZZ'],
    [702, 'AAA'],
  ])('coluna %i é %s', (index, name) => {
    // Excel's base 26 has no zero: after Z comes AA, not BA. The classic mistake here, and it only
    // shows at column 26.
    expect(columnName(index)).toBe(name)
  })

  it('volta do nome para o índice', () => {
    for (const index of [0, 25, 26, 51, 52, 701, 702]) {
      expect(columnIndex(columnName(index))).toBe(index)
    }
  })

  it('recusa nome que não é coluna', () => {
    expect(columnIndex('A1')).toBe(-1)
    expect(columnIndex('')).toBe(-1)
  })
})

describe('referências A1', () => {
  it('monta a referência a partir da posição', () => {
    expect(cellRef(0, 0)).toBe('A1')
    expect(cellRef(9, 27)).toBe('AB10')
  })

  it('lê a referência de volta', () => {
    expect(parseRef('AB10')).toEqual({ row: 9, column: 27 })
    expect(parseRef(' b2 ')).toEqual({ row: 1, column: 1 })
  })

  it.each(['', 'A', '1', 'A0', '1A', 'A-1'])('recusa %o', (ref) => {
    expect(parseRef(ref)).toBeNull()
  })
})

describe('escrita de célula', () => {
  it('guarda e lê de volta', () => {
    const sheet = setCell(createSheet('S'), 2, 1, { value: 42 })

    expect(getCell(sheet, 2, 1)).toEqual({ value: 42 })
    expect(sheet.cells['B3']).toEqual({ value: 42 })
  })

  it('não altera a planilha original', () => {
    const original = createSheet('S')
    setCell(original, 0, 0, { value: 1 })

    expect(original.cells).toEqual({})
  })

  it('remove a célula que ficou sem nada', () => {
    // Clearing an area must not leave thousands of empty entries in the file: it would grow with
    // every clear, without the user understanding why.
    const withValue = setCell(createSheet('S'), 0, 0, { value: 'x' })
    const cleared = setCell(withValue, 0, 0, { value: '' })

    expect(Object.keys(cleared.cells)).toEqual([])
  })

  it('mantém a célula que só tem formatação', () => {
    // A yellow background without text is content: the user painted it on purpose.
    const painted = setCell(createSheet('S'), 0, 0, { style: { background: '#ffff00' } })

    expect(Object.keys(painted.cells)).toEqual(['A1'])
  })

  it('mantém a célula que só tem fórmula', () => {
    const withFormula = setCell(createSheet('S'), 0, 0, { formula: '=SOMA(B1:B9)' })

    expect(Object.keys(withFormula.cells)).toEqual(['A1'])
  })
})
