import { describe, expect, it } from 'vitest'
import { createSheet, getCell, parseRef, setCell, type Sheet, type WorkbookModel } from '../model.js'
import { FormulaError } from './errors.js'
import { recalculate } from './recalc.js'

/** A sheet from a map of reference → value or formula. */
function sheetWith(name: string, entries: Record<string, string | number>): Sheet {
  let sheet = createSheet(name)

  for (const [ref, entry] of Object.entries(entries)) {
    const { row, column } = parseRef(ref)!
    sheet =
      typeof entry === 'string' && entry.startsWith('=')
        ? setCell(sheet, row, column, { formula: entry })
        : setCell(sheet, row, column, { value: entry })
  }

  return sheet
}

function workbookWith(...sheets: Sheet[]): WorkbookModel {
  return { sheets, activeSheet: 0 }
}

function valueOf(workbook: WorkbookModel, ref: string, sheet = 0) {
  const { row, column } = parseRef(ref)!
  return getCell(workbook.sheets[sheet]!, row, column)?.value
}

describe('calculation order', () => {
  it('calculates the chain back to front, not in map order', () => {
    // A1 depends on B1, which depends on C1. Computing in appearance order would give A1 the old
    // value of B1, and the error would only show the second time someone touched the sheet.
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=B1+1', B1: '=C1*2', C1: 5 })))

    expect(valueOf(done, 'C1')).toBe(5)
    expect(valueOf(done, 'B1')).toBe(10)
    expect(valueOf(done, 'A1')).toBe(11)
  })

  it('resolves a long chain without overflowing the stack', () => {
    // A recursive walk would die here, and the user would see the app vanish without any
    // explanation. The map is built directly because `setCell` copies the whole sheet for each
    // cell, and five thousand copies would dominate the test.
    const cells: Record<string, { formula?: string; value?: number }> = { A1: { value: 1 } }
    for (let row = 2; row <= 5000; row++) cells[`A${row}`] = { formula: `=A${row - 1}+1` }

    const done = recalculate(workbookWith({ ...createSheet('P'), cells }))

    expect(valueOf(done, 'A5000')).toBe(5000)
  })

  it('propagates through a range', () => {
    const done = recalculate(
      workbookWith(sheetWith('P', { A1: 1, A2: '=A1*10', A3: '=A2*10', B1: '=SOMA(A1:A3)' })),
    )

    expect(valueOf(done, 'B1')).toBe(111)
  })
})

describe('circular reference', () => {
  it('marks both ends of a direct cycle', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=B1', B1: '=A1' })))

    expect(valueOf(done, 'A1')).toBe(FormulaError.Circular)
    expect(valueOf(done, 'B1')).toBe(FormulaError.Circular)
  })

  it('catches a cell that references itself', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=A1+1' })))

    expect(valueOf(done, 'A1')).toBe(FormulaError.Circular)
  })

  it('catches the cycle through the sum of a range containing the cell itself', () => {
    // The most common way to create a cycle unknowingly: dragging the sum into the column it sums.
    const done = recalculate(workbookWith(sheetWith('P', { A1: 1, A2: 2, A3: '=SOMA(A1:A3)' })))

    expect(valueOf(done, 'A3')).toBe(FormulaError.Circular)
  })

  it('a cell depending on a circular one inherits the error, not a stale value', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=B1', B1: '=A1', C1: '=A1+1' })))

    expect(valueOf(done, 'C1')).toBe(FormulaError.Circular)
  })

  it('the rest of the sheet keeps calculating', () => {
    // A cycle in one corner must not bring down the rest of the calculation.
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=B1', B1: '=A1', D1: 2, D2: '=D1*3' })))

    expect(valueOf(done, 'D2')).toBe(6)
  })
})

describe('across sheets', () => {
  it('reads a cell from another sheet', () => {
    const done = recalculate(
      workbookWith(sheetWith('Resumo', { A1: '=Dados!B2*2' }), sheetWith('Dados', { B2: 21 })),
    )

    expect(valueOf(done, 'A1')).toBe(42)
  })

  it("an unqualified reference points to the formula's own sheet", () => {
    // The trap: if the default were the active tab, the second tab's formula would read the first
    // one's cell and give a plausible, wrong number.
    const done = recalculate(
      workbookWith(sheetWith('Um', { A1: 100, B1: '=A1' }), sheetWith('Dois', { A1: 7, B1: '=A1' })),
    )

    expect(valueOf(done, 'B1', 0)).toBe(100)
    expect(valueOf(done, 'B1', 1)).toBe(7)
  })

  it('ignores case in the sheet name', () => {
    const done = recalculate(
      workbookWith(sheetWith('Resumo', { A1: '=dados!B2' }), sheetWith('Dados', { B2: 9 })),
    )

    expect(valueOf(done, 'A1')).toBe(9)
  })

  it('a sheet that does not exist becomes #REF!', () => {
    const done = recalculate(workbookWith(sheetWith('Resumo', { A1: '=Sumida!B2' })))

    expect(valueOf(done, 'A1')).toBe(FormulaError.Ref)
  })

  it('detects a cycle that crosses sheets', () => {
    const done = recalculate(
      workbookWith(sheetWith('Um', { A1: '=Dois!A1' }), sheetWith('Dois', { A1: '=Um!A1' })),
    )

    expect(valueOf(done, 'A1', 0)).toBe(FormulaError.Circular)
    expect(valueOf(done, 'A1', 1)).toBe(FormulaError.Circular)
  })
})

describe('what recalculation preserves', () => {
  it('keeps the formula next to the value', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: 2, B1: '=A1*3' })))

    expect(getCell(done.sheets[0]!, 0, 1)).toEqual({ formula: '=A1*3', value: 6 })
  })

  it('returns the same workbook when there are no formulas', () => {
    // Identity matters: React redraws the whole grid if the object changes.
    const workbook = workbookWith(sheetWith('P', { A1: 1, B1: 2 }))

    expect(recalculate(workbook)).toBe(workbook)
  })

  it('returns the same workbook when no value changed', () => {
    const workbook = recalculate(workbookWith(sheetWith('P', { A1: 2, B1: '=A1*3' })))

    expect(recalculate(workbook)).toBe(workbook)
  })

  it('a formula over an empty cell is zero and does not vanish from the file', () => {
    // If an empty result erased the cell, the formula would go with it: the sparse map removes
    // cells without value, formula or style.
    const done = recalculate(workbookWith(sheetWith('P', { B1: '=A1' })))

    expect(getCell(done.sheets[0]!, 0, 1)).toEqual({ formula: '=A1', value: 0 })
  })

  it('a formula that does not close does not break recalculation', () => {
    // Only reachable with a hand-edited file; the UI refuses it first.
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=SOMA(', B1: 1, C1: '=B1+1' })))

    expect(valueOf(done, 'A1')).toBe(FormulaError.Value)
    expect(valueOf(done, 'C1')).toBe(2)
  })
})

describe('injected date', () => {
  it('HOJE uses the clock it receives', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=ANO(HOJE())' })), () => new Date(2030, 0, 1))

    expect(valueOf(done, 'A1')).toBe(2030)
  })
})
