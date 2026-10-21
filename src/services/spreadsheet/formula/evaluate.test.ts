import { describe, expect, it } from 'vitest'
import { cellRef } from '../model.js'
import { FormulaError } from './errors.js'
import { evaluate, type EvalContext } from './evaluate.js'
import { parseFormula } from './parse.js'
import type { Scalar } from './values.js'

/** A fake context: a handful of cells and a stopped clock. */
function context(cells: Record<string, Scalar> = {}): EvalContext {
  return {
    valueAt: (ref) => cells[cellRef(ref.row, ref.column)] ?? null,
    now: () => new Date(2026, 7, 15, 12, 0, 0),
  }
}

const calc = (formula: string, cells?: Record<string, Scalar>): Scalar =>
  evaluate(parseFormula(formula), context(cells))

describe('arithmetic', () => {
  it.each([
    ['=1+2', 3],
    ['=10/4', 2.5],
    ['=2*3+4', 10],
    ['=2^10', 1024],
    ['=-2^2', 4],
    ['=10%', 0.1],
    ['=(1+2)*3', 9],
  ])('%s = %s', (formula, expected) => {
    expect(calc(formula)).toBe(expected)
  })

  it('division by zero becomes an error, not infinity', () => {
    expect(calc('=1/0')).toBe(FormulaError.Div0)
  })

  it('overflow becomes #NÚM!', () => {
    expect(calc('=1e308*10')).toBe(FormulaError.Num)
  })

  it('square root of a negative is #NÚM!, not NaN', () => {
    expect(calc('=RAIZ(-1)')).toBe(FormulaError.Num)
  })
})

describe('errors propagate', () => {
  it('run through the whole calculation', () => {
    expect(calc('=SOMA(1;2)+1/0')).toBe(FormulaError.Div0)
  })

  it('the first error is the one that shows', () => {
    expect(calc('=1/0+A1', { A1: FormulaError.NA })).toBe(FormulaError.Div0)
  })

  it('SEERRO intercepts', () => {
    expect(calc('=SEERRO(1/0;"sem valor")')).toBe('sem valor')
    expect(calc('=SEERRO(10;"sem valor")')).toBe(10)
  })

  it('ÉERROS sees the error instead of propagating it', () => {
    // Without special handling, the argument's error would become the result before the function is
    // called, and ÉERROS could never answer TRUE.
    expect(calc('=ÉERROS(1/0)')).toBe(true)
    expect(calc('=ÉERROS(1)')).toBe(false)
  })
})

describe('SE does not evaluate the discarded branch', () => {
  it('does not blow up on a division in the branch not chosen', () => {
    // The case behind lazy evaluation: with A1 at zero, computing 1/A1 would produce #DIV/0! even
    // though the user guarded the formula.
    expect(calc('=SE(A1=0;"";1/A1)', { A1: 0 })).toBe('')
    expect(calc('=SE(A1=0;"";1/A1)', { A1: 4 })).toBe(0.25)
  })

  it('without the third argument returns FALSO, as Excel does', () => {
    expect(calc('=SE(1>2;"sim")')).toBe(false)
  })
})

describe('empty cells', () => {
  it('count as zero in calculations', () => {
    expect(calc('=A1+1')).toBe(1)
  })

  it('are left out of the average', () => {
    // If empty became zero, the average of 10 and 20 with an empty cell between would be 10 instead
    // of 15: wrong, and without any warning.
    expect(calc('=MÉDIA(A1:A3)', { A1: 10, A3: 20 })).toBe(15)
  })

  it('are not counted by CONT.NÚM or CONT.VALORES', () => {
    expect(calc('=CONT.NÚM(A1:A3)', { A1: 10, A3: 20 })).toBe(2)
    expect(calc('=CONT.VALORES(A1:A3)', { A1: 'x', A3: 20 })).toBe(2)
  })

  it('compare equal to both zero and empty text', () => {
    expect(calc('=A1=0')).toBe(true)
    expect(calc('=A1=""')).toBe(true)
  })
})

describe('comparison', () => {
  it('text comparison ignores case', () => {
    expect(calc('="a"="A"')).toBe(true)
  })

  it('different types are never equal', () => {
    // =1="1" is false in Excel. Comparing by conversion would send a text column down a decision
    // path that is not its own.
    expect(calc('=1="1"')).toBe(false)
  })

  it('sorts numbers before text', () => {
    expect(calc('=1<"a"')).toBe(true)
  })
})

describe('text in calculations', () => {
  it('inside a range it is ignored', () => {
    // A column with a text header must still sum.
    expect(calc('=SOMA(A1:A3)', { A1: 'total', A2: 10, A3: 20 })).toBe(30)
  })

  it('passed directly it is converted', () => {
    expect(calc('=SOMA("5";1)')).toBe(6)
  })

  it('text that is not a number becomes #VALOR!', () => {
    expect(calc('=1+"abc"')).toBe(FormulaError.Value)
  })

  it('reads a number in Brazilian format', () => {
    expect(calc('=A1+0', { A1: '1.234,56' })).toBe(1234.56)
  })
})

describe('concatenation', () => {
  it('joins text and number', () => {
    expect(calc('="total: "&A1', { A1: 42 })).toBe('total: 42')
  })

  it('uses a decimal point, which is stable across machines', () => {
    expect(calc('=""&1,5')).toBe('1.5')
  })

  it('empty becomes empty text, not the word "nulo"', () => {
    expect(calc('="x"&A1')).toBe('x')
  })
})

describe('ranges', () => {
  it('sum the whole rectangle', () => {
    expect(calc('=SOMA(A1:B2)', { A1: 1, B1: 2, A2: 3, B2: 4 })).toBe(10)
  })

  it('loose in a calculation are an error', () => {
    expect(calc('=A1:B2+1', { A1: 1 })).toBe(FormulaError.Value)
  })

  it('work when written backwards', () => {
    expect(calc('=SOMA(B2:A1)', { A1: 1, B1: 2, A2: 3, B2: 4 })).toBe(10)
  })
})

describe('rounding', () => {
  it('halves round away from zero, for both signs', () => {
    // JavaScript's Math.round would take -2.5 to -2, and the one-cent difference would only show at
    // month end.
    expect(calc('=ARRED(2,5;0)')).toBe(3)
    expect(calc('=ARRED(-2,5;0)')).toBe(-3)
  })

  it('does not slip on binary representation', () => {
    expect(calc('=ARRED(1,005;2)')).toBe(1.01)
  })

  it('accepts negative places', () => {
    expect(calc('=ARRED(1234;-2)')).toBe(1200)
  })

  it('INT really rounds down, TRUNCAR cuts', () => {
    expect(calc('=INT(-2,5)')).toBe(-3)
    expect(calc('=TRUNCAR(-2,5)')).toBe(-2)
  })

  it("RESTO follows the divisor's sign, as in Excel", () => {
    // JavaScript's % operator would give -1 here.
    expect(calc('=RESTO(-1;3)')).toBe(2)
  })
})

describe('aggregation with criteria', () => {
  const table = { A1: 'sul', B1: 10, A2: 'norte', B2: 20, A3: 'sul', B3: 30 }

  it('SOMASE with text', () => {
    expect(calc('=SOMASE(A1:A3;"sul";B1:B3)', table)).toBe(40)
  })

  it('SOMASE with a comparator', () => {
    expect(calc('=SOMASE(B1:B3;">15")', table)).toBe(50)
  })

  it('CONT.SE with a wildcard', () => {
    expect(calc('=CONT.SE(A1:A3;"s*")', table)).toBe(2)
  })

  it('CONT.SE does not count empty cells with a not-equal criterion', () => {
    // Otherwise a column of ten thousand blank rows would count ten thousand.
    expect(calc('=CONT.SE(A1:A9;"<>norte")', table)).toBe(2)
  })
})

describe('lookup', () => {
  const table = { A1: 1, B1: 'um', A2: 5, B2: 'cinco', A3: 9, B3: 'nove' }

  it('exact PROCV', () => {
    expect(calc('=PROCV(5;A1:B3;2;FALSO)', table)).toBe('cinco')
  })

  it('exact PROCV with no match returns #N/D', () => {
    expect(calc('=PROCV(7;A1:B3;2;FALSO)', table)).toBe(FormulaError.NA)
  })

  it('approximate PROCV is the default, as in Excel', () => {
    // Without the fourth argument Excel searches approximately. It is a bad default, but changing
    // it would make the same spreadsheet give different results in each.
    expect(calc('=PROCV(7;A1:B3;2)', table)).toBe('cinco')
  })

  it('a column outside the range is #REF!', () => {
    expect(calc('=PROCV(5;A1:B3;3;FALSO)', table)).toBe(FormulaError.Ref)
  })

  it('CORRESP and ÍNDICE work together', () => {
    expect(calc('=ÍNDICE(B1:B3;CORRESP(9;A1:A3;0))', table)).toBe('nove')
  })
})

describe('text', () => {
  it.each([
    ['=ESQUERDA("planilha";4)', 'plan'],
    ['=DIREITA("planilha";4)', 'ilha'],
    ['=EXT.TEXTO("planilha";3;4)', 'anil'],
    ['=NÚM.CARACT("planilha")', 8],
    ['=MAIÚSCULA("olá")', 'OLÁ'],
    ['=ARRUMAR("  a   b  ")', 'a b'],
    ['=SUBSTITUIR("a-b-c";"-";"/")', 'a/b/c'],
    ['=CONCATENAR("a";1;VERDADEIRO)', 'a1VERDADEIRO'],
    ['=PROCURAR("IL";"planilha")', 5],
  ])('%s = %s', (formula, expected) => {
    expect(calc(formula)).toBe(expected)
  })

  it('LOCALIZAR is case-sensitive and PROCURAR is not', () => {
    expect(calc('=PROCURAR("IL";"planilha")')).toBe(5)
    expect(calc('=SEERRO(LOCALIZAR("IL";"planilha");"não achou")')).toBe('não achou')
  })

  it('counts an emoji as one character', () => {
    expect(calc('=NÚM.CARACT("a🙂")')).toBe(2)
  })
})

describe('dates', () => {
  it('HOJE comes from the context, not the clock', () => {
    // Without injecting "now", every date test would break the next day.
    expect(calc('=ANO(HOJE())')).toBe(2026)
    expect(calc('=MÊS(HOJE())')).toBe(8)
    expect(calc('=DIA(HOJE())')).toBe(15)
  })

  it('the difference between dates is a subtraction', () => {
    expect(calc('=DATA(2026;8;15)-DATA(2026;8;1)')).toBe(14)
  })

  it('month thirteen is January of the next year', () => {
    expect(calc('=ANO(DATA(2026;13;1))')).toBe(2027)
  })
})

describe('logic', () => {
  it('E and OU do not stop at the first result', () => {
    // Excel evaluates every argument: stopping early would return FALSE here and #DIV/0! there, for
    // the same spreadsheet.
    expect(calc('=E(FALSO;1/0)')).toBe(FormulaError.Div0)
    expect(calc('=OU(VERDADEIRO;1/0)')).toBe(FormulaError.Div0)
  })

  it('combine normally', () => {
    expect(calc('=E(1>0;2>1)')).toBe(true)
    expect(calc('=OU(1>2;2>3)')).toBe(false)
    expect(calc('=NÃO(1>0)')).toBe(false)
  })
})

describe('names', () => {
  it('an unknown function becomes #NOME?', () => {
    expect(calc('=INVENTADA(1)')).toBe(FormulaError.Name)
  })

  it('English and Portuguese names are equivalent', () => {
    expect(calc('=SUM(1;2)')).toBe(calc('=SOMA(1;2)'))
    expect(calc('=IF(1>0;"a";"b")')).toBe(calc('=SE(1>0;"a";"b")'))
  })

  it('wrong number of arguments becomes #VALOR!', () => {
    expect(calc('=ABS(1;2)')).toBe(FormulaError.Value)
  })
})
