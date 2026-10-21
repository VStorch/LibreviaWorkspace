import { describe, expect, it } from 'vitest'
import { fromXlsxFormula, toXlsxFormula } from './interop.js'
import { parseFormula } from './parse.js'

describe('fromXlsxFormula', () => {
  it.each([
    ['=SUM(A1,B1)', '=SOMA(A1;B1)'],
    ['=AVERAGE(D2:D9)', '=MÉDIA(D2:D9)'],
    ['=IF(A1>0,"sim","não")', '=SE(A1>0;"sim";"não")'],
    ['=IFERROR(A1/B1,0)', '=SEERRO(A1/B1;0)'],
    ['=VLOOKUP(A1,Tabela!A:B,2,FALSE)', '=PROCV(A1;Tabela!A:B;2;FALSE)'],
    ['=ROUND(A1*1.5,2)', '=ARRED(A1*1,5;2)'],
    ['=COUNTIF(A1:A9,">10")', '=CONT.SE(A1:A9;">10")'],
    ['=CONCATENATE(A1," ",B1)', '=CONCATENAR(A1;" ";B1)'],
    ['=FIND("a",A1)', '=PROCURAR("a";A1)'],
    ['=SEARCH("a",A1)', '=LOCALIZAR("a";A1)'],
  ])('translates %s', (file, app) => {
    expect(fromXlsxFormula(file)).toBe(app)
  })

  it('leaves commas inside text alone', () => {
    expect(fromXlsxFormula('=IF(A1,"um, dois","três")')).toBe('=SE(A1;"um, dois";"três")')
  })

  it('leaves commas inside a sheet name alone', () => {
    expect(fromXlsxFormula("='Vendas, filial'!A1")).toBe("='Vendas, filial'!A1")
  })

  it('leaves the literal array as is', () => {
    // The app does not compute arrays. Translating what it does not understand would spoil a file
    // that would stay fine if left alone.
    expect(fromXlsxFormula('=SUM({1,2;3,4})')).toBe('=SOMA({1,2;3,4})')
  })

  it('translates an error written in the formula', () => {
    expect(fromXlsxFormula('=IFERROR(A1,#N/A)')).toBe('=SEERRO(A1;#N/D)')
  })

  it('does not confuse a cell with a function call', () => {
    // `SOMA` without an attached parenthesis is a reference, not a call.
    expect(fromXlsxFormula('=SUM(SUM1,B1)')).toBe('=SOMA(SUM1;B1)')
  })

  it('leaves intact a function the engine does not know', () => {
    expect(fromXlsxFormula('=_xlfn.XLOOKUP(A1,B:B,C:C)')).toBe('=_xlfn.XLOOKUP(A1;B:B;C:C)')
  })

  it('produces a formula the parser accepts', () => {
    expect(() => parseFormula(fromXlsxFormula('=IF(SUM(A1:A3)>1.5,ROUND(B1,2),0)'))).not.toThrow()
  })
})

describe('toXlsxFormula', () => {
  it.each([
    ['=SOMA(A1;B1)', '=SUM(A1,B1)'],
    ['=MÉDIA(D2:D9)', '=AVERAGE(D2:D9)'],
    ['=SE(A1>0;"sim";"não")', '=IF(A1>0,"sim","não")'],
    ['=ARRED(A1*1,5;2)', '=ROUND(A1*1.5,2)'],
    ['=CONT.NÚM(A1:A9)', '=COUNT(A1:A9)'],
    ['=NÚM.CARACT(A1)', '=LEN(A1)'],
    ['=ÉERROS(A1)', '=ISERROR(A1)'],
    ['=SENÃODISP(A1;0)', '=IFNA(A1,0)'],
  ])('translates %s', (app, file) => {
    expect(toXlsxFormula(app)).toBe(file)
  })

  it('accepts an English name already typed by the user', () => {
    expect(toXlsxFormula('=SUM(A1;B1)')).toBe('=SUM(A1,B1)')
  })

  it('translates an error written in the formula', () => {
    expect(toXlsxFormula('=SEERRO(A1;#N/D)')).toBe('=IFERROR(A1,#N/A)')
  })

  it('leaves semicolons inside text alone', () => {
    expect(toXlsxFormula('=CONCATENAR("a;b";A1)')).toBe('=CONCATENATE("a;b",A1)')
  })
})

describe('round trip', () => {
  // The surgical save compares the outgoing formula with the one in the file. If the round trip
  // changed anything, opening and saving without editing would rewrite every cell and erase what
  // the model does not represent.
  it.each([
    '=SUM(D2:D3)',
    '=B2*C2',
    '=IF(A1>0,SUM(A1:A9),0)',
    '=ROUND(A1*1.5,2)',
    "='Plan 2'!A1+Plan1!B2",
    '=SUM(A1:A3)+_xlfn.XLOOKUP(A1,B:B,C:C)',
    '=IFERROR(VLOOKUP(A1,T!A:B,2,FALSE),"—")',
    '=SUM({1,2;3,4})',
    '=A1&", "&B1',
  ])('%s comes back identical', (formula) => {
    expect(toXlsxFormula(fromXlsxFormula(formula))).toBe(formula)
  })

  it('preserves spaces and case of what it does not translate', () => {
    expect(toXlsxFormula(fromXlsxFormula('=SUM( a1 , b1 )'))).toBe('=SUM( a1 , b1 )')
  })
})
