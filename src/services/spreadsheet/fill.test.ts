import { describe, expect, it } from 'vitest'
import { fillRange } from './fill.js'
import { singleCell } from './edit.js'
import { CellFormat, createSheet, getCell, setCell, type Sheet } from './model.js'

function sheetWith(cells: Array<[number, number, Parameters<typeof setCell>[3]]>): Sheet {
  return cells.reduce((sheet, [row, column, cell]) => setCell(sheet, row, column, cell), createSheet('Plan1'))
}

describe('fillRange', () => {
  it('desloca a fórmula, em vez de repetir o resultado', () => {
    // The difference between a spreadsheet and a table: whoever drags `=B1*C1` expects `=B2*C2`,
    // not the first row's number three times.
    const sheet = sheetWith([[0, 3, { formula: '=B1*C1', value: 6 }]])

    const filled = fillRange(sheet, singleCell(0, 3), { fromRow: 0, toRow: 2, fromColumn: 3, toColumn: 3 })

    expect(getCell(filled, 1, 3)?.formula).toBe('=B2*C2')
    expect(getCell(filled, 2, 3)?.formula).toBe('=B3*C3')
  })

  it('não leva o valor calculado da origem', () => {
    // The value belongs to the source position. Recalculation fills the new cell's, since it knows
    // the right order.
    const sheet = sheetWith([[0, 0, { formula: '=1+1', value: 2 }]])

    const filled = fillRange(sheet, singleCell(0, 0), { fromRow: 0, toRow: 1, fromColumn: 0, toColumn: 0 })

    expect(getCell(filled, 1, 0)?.value).toBeUndefined()
  })

  it('a referência com $ não acompanha', () => {
    const sheet = sheetWith([[0, 1, { formula: '=A1*$C$1' }]])

    const filled = fillRange(sheet, singleCell(0, 1), { fromRow: 0, toRow: 1, fromColumn: 1, toColumn: 1 })

    expect(getCell(filled, 1, 1)?.formula).toBe('=A2*$C$1')
  })

  it('preenche para o lado deslocando a coluna', () => {
    const sheet = sheetWith([[0, 0, { formula: '=A2' }]])

    const filled = fillRange(sheet, singleCell(0, 0), { fromRow: 0, toRow: 0, fromColumn: 0, toColumn: 2 })

    expect(getCell(filled, 0, 1)?.formula).toBe('=B2')
    expect(getCell(filled, 0, 2)?.formula).toBe('=C2')
  })

  it('leva a formatação junto', () => {
    // In Excel the handle carries the format. Leaving it behind would turn a filled currency column
    // into half a column of raw numbers.
    const sheet = sheetWith([[0, 0, { value: 10, style: { format: CellFormat.Currency, bold: true } }]])

    const filled = fillRange(sheet, singleCell(0, 0), { fromRow: 0, toRow: 1, fromColumn: 0, toColumn: 0 })

    expect(getCell(filled, 1, 0)?.style).toEqual({ format: CellFormat.Currency, bold: true })
    expect(getCell(filled, 1, 0)?.value).toBe(10)
  })

  it('não reescreve a origem', () => {
    const sheet = sheetWith([[0, 0, { formula: '=X1', value: 5 }]])

    const filled = fillRange(sheet, singleCell(0, 0), { fromRow: 0, toRow: 3, fromColumn: 0, toColumn: 0 })

    expect(getCell(filled, 0, 0)).toEqual({ formula: '=X1', value: 5 })
  })

  it('repete o padrão ciclicamente', () => {
    // Two source rows dragged over six fill the pattern three times, as in Excel.
    const sheet = sheetWith([
      [0, 0, { value: 'par' }],
      [1, 0, { value: 'ímpar' }],
    ])

    const filled = fillRange(
      sheet,
      { fromRow: 0, toRow: 1, fromColumn: 0, toColumn: 0 },
      { fromRow: 0, toRow: 5, fromColumn: 0, toColumn: 0 },
    )

    expect(getCell(filled, 2, 0)?.value).toBe('par')
    expect(getCell(filled, 3, 0)?.value).toBe('ímpar')
    expect(getCell(filled, 4, 0)?.value).toBe('par')
  })

  it('preencher para cima também funciona', () => {
    // The modulo of a negative number is negative in JavaScript, and without a sign fix it would
    // point outside the source.
    const sheet = sheetWith([[3, 0, { formula: '=B4' }]])

    const filled = fillRange(sheet, singleCell(3, 0), { fromRow: 1, toRow: 3, fromColumn: 0, toColumn: 0 })

    expect(getCell(filled, 2, 0)?.formula).toBe('=B3')
    expect(getCell(filled, 1, 0)?.formula).toBe('=B2')
  })

  it('origem vazia limpa o destino', () => {
    // Dragging an empty cell over data erases it, as in Excel.
    const sheet = sheetWith([
      [0, 0, {}],
      [1, 0, { value: 'apagar' }],
    ])

    const filled = fillRange(sheet, singleCell(0, 0), { fromRow: 0, toRow: 1, fromColumn: 0, toColumn: 0 })

    expect(getCell(filled, 1, 0)).toBeUndefined()
  })

  it('a fórmula que sairia da planilha vira #REF!', () => {
    const sheet = sheetWith([[2, 0, { formula: '=A1' }]])

    const filled = fillRange(sheet, singleCell(2, 0), { fromRow: 0, toRow: 2, fromColumn: 0, toColumn: 0 })

    expect(getCell(filled, 0, 0)?.formula).toContain('#REF!')
  })
})
