import { describe, expect, it } from 'vitest'
import { recalculate } from './formula/recalc.js'
import { createSheet, getCell, setCell, type Sheet, type WorkbookModel } from './model.js'
import { Language } from '@shared/i18n/index.js'
import { applyStructuralChange, isNameTaken, nextSheetName } from './structure.js'

/** Column A with 1, 2, 3 and their sum in C1. */
function makeSheet(name = 'Plan1'): Sheet {
  let sheet = createSheet(name)
  sheet = setCell(sheet, 0, 0, { value: 1 })
  sheet = setCell(sheet, 1, 0, { value: 2 })
  sheet = setCell(sheet, 2, 0, { value: 3 })
  sheet = setCell(sheet, 0, 2, { formula: '=SOMA(A1:A3)', value: 6 })
  return sheet
}

const makeWorkbook = (...sheets: Sheet[]): WorkbookModel => ({ sheets, activeSheet: 0 })

describe('inserting a row', () => {
  it('stretches the formula range along with the data', () => {
    // Without the adjustment the data moved down and the formula kept summing the old place, giving
    // a wrong total without warning.
    const depois = applyStructuralChange(makeWorkbook(makeSheet()), 0, {
      kind: 'insertRows',
      at: 1,
      count: 1,
    })

    expect(getCell(depois.sheets[0]!, 0, 2)?.formula).toBe('=SOMA(A1:A4)')
  })

  it('and the total stays correct after recalculating', () => {
    let currentWorkbook = applyStructuralChange(makeWorkbook(makeSheet()), 0, {
      kind: 'insertRows',
      at: 1,
      count: 1,
    })
    currentWorkbook = recalculate(currentWorkbook)

    expect(getCell(currentWorkbook.sheets[0]!, 0, 2)?.value).toBe(6)
  })

  it('filling the new row, the total follows', () => {
    let currentWorkbook = applyStructuralChange(makeWorkbook(makeSheet()), 0, {
      kind: 'insertRows',
      at: 1,
      count: 1,
    })
    const comValor = setCell(currentWorkbook.sheets[0]!, 1, 0, { value: 10 })
    currentWorkbook = recalculate({ ...currentWorkbook, sheets: [comValor] })

    expect(getCell(currentWorkbook.sheets[0]!, 0, 2)?.value).toBe(16)
  })
})

describe('deleting a row', () => {
  // Row 2 is deleted: row 1 takes along the formula's own cell, which lives in it.
  it('shrinks the range instead of breaking the formula', () => {
    const depois = applyStructuralChange(makeWorkbook(makeSheet()), 0, {
      kind: 'deleteRows',
      at: 1,
      count: 1,
    })

    expect(getCell(depois.sheets[0]!, 0, 2)?.formula).toBe('=SOMA(A1:A2)')
  })

  it('the total matches what is left', () => {
    const depois = recalculate(
      applyStructuralChange(makeWorkbook(makeSheet()), 0, { kind: 'deleteRows', at: 1, count: 1 }),
    )

    expect(getCell(depois.sheets[0]!, 0, 2)?.value).toBe(4)
  })

  it('deleting the referenced cell gives #REF!', () => {
    let sheet = createSheet('Plan1')
    sheet = setCell(sheet, 4, 0, { value: 9 })
    sheet = setCell(sheet, 0, 1, { formula: '=A5' })

    const depois = applyStructuralChange(makeWorkbook(sheet), 0, { kind: 'deleteRows', at: 4, count: 1 })

    expect(getCell(depois.sheets[0]!, 0, 1)?.formula).toBe('=#REF!')
  })
})

describe('columns', () => {
  it('inserting pushes the reference to the right', () => {
    const depois = applyStructuralChange(makeWorkbook(makeSheet()), 0, {
      kind: 'insertColumns',
      at: 0,
      count: 1,
    })

    expect(getCell(depois.sheets[0]!, 0, 3)?.formula).toBe('=SOMA(B1:B3)')
  })
})

describe('across sheets', () => {
  it('a formula on another sheet follows the row inserted here', () => {
    // Why the operation belongs to the workbook and not the sheet: otherwise =Dados!A5 would keep
    // pointing where the data no longer is.
    let resumo = createSheet('Resumo')
    resumo = setCell(resumo, 0, 0, { formula: '=Dados!A5' })

    const depois = applyStructuralChange(makeWorkbook(makeSheet('Dados'), resumo), 0, {
      kind: 'insertRows',
      at: 0,
      count: 1,
    })

    expect(getCell(depois.sheets[1]!, 0, 0)?.formula).toBe('=Dados!A6')
  })

  it('an unqualified reference on another sheet does not move', () => {
    // =A5 on "Resumo" points to "Resumo", not to "Dados".
    let resumo = createSheet('Resumo')
    resumo = setCell(resumo, 0, 0, { formula: '=A5' })

    const depois = applyStructuralChange(makeWorkbook(makeSheet('Dados'), resumo), 0, {
      kind: 'insertRows',
      at: 0,
      count: 1,
    })

    expect(getCell(depois.sheets[1]!, 0, 0)?.formula).toBe('=A5')
  })
})

describe('identity', () => {
  it('a sheet without formulas comes back as the same object', () => {
    const semFormula = setCell(createSheet('Outra'), 0, 0, { value: 1 })
    const antes = makeWorkbook(makeSheet(), semFormula)
    const depois = applyStructuralChange(antes, 0, { kind: 'insertRows', at: 0, count: 1 })

    expect(depois.sheets[1]).toBe(semFormula)
  })

  it('an operation with no effect returns the workbook untouched', () => {
    const antes = makeWorkbook(makeSheet())

    expect(applyStructuralChange(antes, 0, { kind: 'deleteRows', at: 5000, count: 1 })).toBe(antes)
  })
})

describe('next sheet name', () => {
  it('follows the count when the names are the defaults', () => {
    expect(nextSheetName(makeWorkbook(makeSheet('Planilha1')))).toBe('Planilha2')
  })

  it('skips a name already in use instead of repeating it', () => {
    // Someone who deleted Planilha2 and created another would have two with the same name, and
    // `=Planilha2!A1` would no longer have a single target.
    const workbook = makeWorkbook(makeSheet('Planilha1'), makeSheet('Planilha3'))
    expect(nextSheetName(workbook)).toBe('Planilha4')
  })

  it('names the sheet in the interface language', () => {
    const workbook = makeWorkbook(makeSheet('Sheet1'), makeSheet('Sheet3'))
    expect(nextSheetName(workbook, Language.English)).toBe('Sheet4')
  })

  it('recognizes a name taken by another sheet, not by the sheet itself', () => {
    const atual = makeWorkbook(makeSheet('Dados'), makeSheet('Resumo'))
    expect(isNameTaken(atual, 'Dados', 1)).toBe(true)
    expect(isNameTaken(atual, 'Dados', 0)).toBe(false)
  })
})
