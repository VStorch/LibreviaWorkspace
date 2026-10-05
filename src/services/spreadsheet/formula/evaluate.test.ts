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

describe('aritmética', () => {
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

  it('divisão por zero vira erro, não infinito', () => {
    expect(calc('=1/0')).toBe(FormulaError.Div0)
  })

  it('estouro vira #NÚM!', () => {
    expect(calc('=1e308*10')).toBe(FormulaError.Num)
  })

  it('raiz de negativo é #NÚM!, não NaN', () => {
    expect(calc('=RAIZ(-1)')).toBe(FormulaError.Num)
  })
})

describe('erros se propagam', () => {
  it('atravessam a conta inteira', () => {
    expect(calc('=SOMA(1;2)+1/0')).toBe(FormulaError.Div0)
  })

  it('o primeiro erro é o que aparece', () => {
    expect(calc('=1/0+A1', { A1: FormulaError.NA })).toBe(FormulaError.Div0)
  })

  it('SEERRO intercepta', () => {
    expect(calc('=SEERRO(1/0;"sem valor")')).toBe('sem valor')
    expect(calc('=SEERRO(10;"sem valor")')).toBe(10)
  })

  it('ÉERROS enxerga o erro em vez de propagá-lo', () => {
    // Without special handling, the argument's error would become the result before the function is
    // called, and ÉERROS could never answer TRUE.
    expect(calc('=ÉERROS(1/0)')).toBe(true)
    expect(calc('=ÉERROS(1)')).toBe(false)
  })
})

describe('SE não avalia o ramo descartado', () => {
  it('não estoura na divisão do ramo não escolhido', () => {
    // The case behind lazy evaluation: with A1 at zero, computing 1/A1 would produce #DIV/0! even
    // though the user guarded the formula.
    expect(calc('=SE(A1=0;"";1/A1)', { A1: 0 })).toBe('')
    expect(calc('=SE(A1=0;"";1/A1)', { A1: 4 })).toBe(0.25)
  })

  it('sem o terceiro argumento devolve FALSO, como o Excel', () => {
    expect(calc('=SE(1>2;"sim")')).toBe(false)
  })
})

describe('células vazias', () => {
  it('valem zero na conta', () => {
    expect(calc('=A1+1')).toBe(1)
  })

  it('não entram na média', () => {
    // If empty became zero, the average of 10 and 20 with an empty cell between would be 10 instead
    // of 15: wrong, and without any warning.
    expect(calc('=MÉDIA(A1:A3)', { A1: 10, A3: 20 })).toBe(15)
  })

  it('não são contadas por CONT.NÚM nem por CONT.VALORES', () => {
    expect(calc('=CONT.NÚM(A1:A3)', { A1: 10, A3: 20 })).toBe(2)
    expect(calc('=CONT.VALORES(A1:A3)', { A1: 'x', A3: 20 })).toBe(2)
  })

  it('comparam igual tanto a zero quanto a texto vazio', () => {
    expect(calc('=A1=0')).toBe(true)
    expect(calc('=A1=""')).toBe(true)
  })
})

describe('comparação', () => {
  it('texto não diferencia maiúsculas', () => {
    expect(calc('="a"="A"')).toBe(true)
  })

  it('tipos diferentes nunca são iguais', () => {
    // =1="1" is false in Excel. Comparing by conversion would send a text column down a decision
    // path that is not its own.
    expect(calc('=1="1"')).toBe(false)
  })

  it('ordena número antes de texto', () => {
    expect(calc('=1<"a"')).toBe(true)
  })
})

describe('texto em conta', () => {
  it('dentro de intervalo é ignorado', () => {
    // A column with a text header must still sum.
    expect(calc('=SOMA(A1:A3)', { A1: 'total', A2: 10, A3: 20 })).toBe(30)
  })

  it('passado direto é convertido', () => {
    expect(calc('=SOMA("5";1)')).toBe(6)
  })

  it('texto que não é número vira #VALOR!', () => {
    expect(calc('=1+"abc"')).toBe(FormulaError.Value)
  })

  it('lê número no formato brasileiro', () => {
    expect(calc('=A1+0', { A1: '1.234,56' })).toBe(1234.56)
  })
})

describe('concatenação', () => {
  it('junta texto e número', () => {
    expect(calc('="total: "&A1', { A1: 42 })).toBe('total: 42')
  })

  it('usa ponto decimal, que é estável entre máquinas', () => {
    expect(calc('=""&1,5')).toBe('1.5')
  })

  it('vazio vira texto vazio, e não a palavra "nulo"', () => {
    expect(calc('="x"&A1')).toBe('x')
  })
})

describe('intervalos', () => {
  it('somam o retângulo inteiro', () => {
    expect(calc('=SOMA(A1:B2)', { A1: 1, B1: 2, A2: 3, B2: 4 })).toBe(10)
  })

  it('soltos numa conta são erro', () => {
    expect(calc('=A1:B2+1', { A1: 1 })).toBe(FormulaError.Value)
  })

  it('funcionam escritos ao contrário', () => {
    expect(calc('=SOMA(B2:A1)', { A1: 1, B1: 2, A2: 3, B2: 4 })).toBe(10)
  })
})

describe('arredondamento', () => {
  it('meio sobe para longe do zero, nos dois sinais', () => {
    // JavaScript's Math.round would take -2.5 to -2, and the one-cent difference would only show at
    // month end.
    expect(calc('=ARRED(2,5;0)')).toBe(3)
    expect(calc('=ARRED(-2,5;0)')).toBe(-3)
  })

  it('não escorrega no binário', () => {
    expect(calc('=ARRED(1,005;2)')).toBe(1.01)
  })

  it('aceita casas negativas', () => {
    expect(calc('=ARRED(1234;-2)')).toBe(1200)
  })

  it('INT desce de verdade, TRUNCAR corta', () => {
    expect(calc('=INT(-2,5)')).toBe(-3)
    expect(calc('=TRUNCAR(-2,5)')).toBe(-2)
  })

  it('RESTO acompanha o sinal do divisor, como no Excel', () => {
    // JavaScript's % operator would give -1 here.
    expect(calc('=RESTO(-1;3)')).toBe(2)
  })
})

describe('agregação com critério', () => {
  const tabela = { A1: 'sul', B1: 10, A2: 'norte', B2: 20, A3: 'sul', B3: 30 }

  it('SOMASE com texto', () => {
    expect(calc('=SOMASE(A1:A3;"sul";B1:B3)', tabela)).toBe(40)
  })

  it('SOMASE com comparador', () => {
    expect(calc('=SOMASE(B1:B3;">15")', tabela)).toBe(50)
  })

  it('CONT.SE com curinga', () => {
    expect(calc('=CONT.SE(A1:A3;"s*")', tabela)).toBe(2)
  })

  it('CONT.SE não conta as vazias com critério de diferença', () => {
    // Otherwise a column of ten thousand blank rows would count ten thousand.
    expect(calc('=CONT.SE(A1:A9;"<>norte")', tabela)).toBe(2)
  })
})

describe('procura', () => {
  const tabela = { A1: 1, B1: 'um', A2: 5, B2: 'cinco', A3: 9, B3: 'nove' }

  it('PROCV exato', () => {
    expect(calc('=PROCV(5;A1:B3;2;FALSO)', tabela)).toBe('cinco')
  })

  it('PROCV exato sem achar devolve #N/D', () => {
    expect(calc('=PROCV(7;A1:B3;2;FALSO)', tabela)).toBe(FormulaError.NA)
  })

  it('PROCV aproximado é o padrão, como no Excel', () => {
    // Without the fourth argument Excel searches approximately. It is a bad default, but changing
    // it would make the same spreadsheet give different results in each.
    expect(calc('=PROCV(7;A1:B3;2)', tabela)).toBe('cinco')
  })

  it('coluna fora do intervalo é #REF!', () => {
    expect(calc('=PROCV(5;A1:B3;3;FALSO)', tabela)).toBe(FormulaError.Ref)
  })

  it('CORRESP e ÍNDICE trabalham juntos', () => {
    expect(calc('=ÍNDICE(B1:B3;CORRESP(9;A1:A3;0))', tabela)).toBe('nove')
  })
})

describe('texto', () => {
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

  it('LOCALIZAR diferencia maiúsculas e PROCURAR não', () => {
    expect(calc('=PROCURAR("IL";"planilha")')).toBe(5)
    expect(calc('=SEERRO(LOCALIZAR("IL";"planilha");"não achou")')).toBe('não achou')
  })

  it('conta emoji como um caractere', () => {
    expect(calc('=NÚM.CARACT("a🙂")')).toBe(2)
  })
})

describe('datas', () => {
  it('HOJE vem do contexto, e não do relógio', () => {
    // Without injecting "now", every date test would break the next day.
    expect(calc('=ANO(HOJE())')).toBe(2026)
    expect(calc('=MÊS(HOJE())')).toBe(8)
    expect(calc('=DIA(HOJE())')).toBe(15)
  })

  it('a diferença entre datas é uma subtração', () => {
    expect(calc('=DATA(2026;8;15)-DATA(2026;8;1)')).toBe(14)
  })

  it('mês treze é janeiro do ano seguinte', () => {
    expect(calc('=ANO(DATA(2026;13;1))')).toBe(2027)
  })
})

describe('lógica', () => {
  it('E e OU não param no primeiro resultado', () => {
    // Excel evaluates every argument: stopping early would return FALSE here and #DIV/0! there, for
    // the same spreadsheet.
    expect(calc('=E(FALSO;1/0)')).toBe(FormulaError.Div0)
    expect(calc('=OU(VERDADEIRO;1/0)')).toBe(FormulaError.Div0)
  })

  it('combinam normalmente', () => {
    expect(calc('=E(1>0;2>1)')).toBe(true)
    expect(calc('=OU(1>2;2>3)')).toBe(false)
    expect(calc('=NÃO(1>0)')).toBe(false)
  })
})

describe('nomes', () => {
  it('função desconhecida vira #NOME?', () => {
    expect(calc('=INVENTADA(1)')).toBe(FormulaError.Name)
  })

  it('inglês e português dão no mesmo', () => {
    expect(calc('=SUM(1;2)')).toBe(calc('=SOMA(1;2)'))
    expect(calc('=IF(1>0;"a";"b")')).toBe(calc('=SE(1>0;"a";"b")'))
  })

  it('número errado de argumentos vira #VALOR!', () => {
    expect(calc('=ABS(1;2)')).toBe(FormulaError.Value)
  })
})
