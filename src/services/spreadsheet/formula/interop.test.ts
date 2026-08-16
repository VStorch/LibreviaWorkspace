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
  ])('traduz %s', (file, app) => {
    expect(fromXlsxFormula(file)).toBe(app)
  })

  it('não mexe em vírgula dentro de texto', () => {
    expect(fromXlsxFormula('=IF(A1,"um, dois","três")')).toBe('=SE(A1;"um, dois";"três")')
  })

  it('não mexe em vírgula dentro de nome de planilha', () => {
    expect(fromXlsxFormula("='Vendas, filial'!A1")).toBe("='Vendas, filial'!A1")
  })

  it('deixa a matriz literal como está', () => {
    // O aplicativo não calcula matriz. Traduzir o que não se entende estragaria
    // um arquivo que continuaria bom se ficasse quieto.
    expect(fromXlsxFormula('=SUM({1,2;3,4})')).toBe('=SOMA({1,2;3,4})')
  })

  it('traduz erro escrito na fórmula', () => {
    expect(fromXlsxFormula('=IFERROR(A1,#N/A)')).toBe('=SEERRO(A1;#N/D)')
  })

  it('não confunde célula chamada como função', () => {
    // `SOMA` sem parêntese colado é referência, não chamada.
    expect(fromXlsxFormula('=SUM(SUM1,B1)')).toBe('=SOMA(SUM1;B1)')
  })

  it('deixa intacta a função que o motor não conhece', () => {
    expect(fromXlsxFormula('=_xlfn.XLOOKUP(A1,B:B,C:C)')).toBe('=_xlfn.XLOOKUP(A1;B:B;C:C)')
  })

  it('entrega fórmula que o analisador aceita', () => {
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
  ])('traduz %s', (app, file) => {
    expect(toXlsxFormula(app)).toBe(file)
  })

  it('aceita o nome em inglês já digitado pelo usuário', () => {
    expect(toXlsxFormula('=SUM(A1;B1)')).toBe('=SUM(A1,B1)')
  })

  it('traduz erro escrito na fórmula', () => {
    expect(toXlsxFormula('=SEERRO(A1;#N/D)')).toBe('=IFERROR(A1,#N/A)')
  })

  it('não mexe em ponto e vírgula dentro de texto', () => {
    expect(toXlsxFormula('=CONCATENAR("a;b";A1)')).toBe('=CONCATENATE("a;b",A1)')
  })
})

describe('ida e volta', () => {
  // A gravação cirúrgica compara a fórmula que sai com a que estava no arquivo.
  // Se a ida e volta mudasse qualquer coisa, abrir e salvar sem editar
  // reescreveria todas as células — e apagaria o que o modelo não representa.
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
  ])('%s volta idêntica', (formula) => {
    expect(toXlsxFormula(fromXlsxFormula(formula))).toBe(formula)
  })

  it('preserva espaços e maiúsculas do que não traduz', () => {
    expect(toXlsxFormula(fromXlsxFormula('=SUM( a1 , b1 )'))).toBe('=SUM( a1 , b1 )')
  })
})
