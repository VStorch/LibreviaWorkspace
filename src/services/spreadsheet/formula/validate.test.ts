import { describe, expect, it } from 'vitest'
import { Language } from '@shared/i18n/index.js'
import { checkFormula } from './validate.js'

describe('valid formula', () => {
  it.each(['=1+1', '=SOMA(A1:A9)', '=SE(A1>0;"sim";"não")', '=HOJE()', '=SOMA(1;2;3;4;5)'])(
    '%s passes',
    (formula) => {
      expect(checkFormula(formula)).toBeNull()
    },
  )
})

describe('syntax problems', () => {
  it('points to the position', () => {
    const problem = checkFormula('=SOMA(A1')

    expect(problem?.message).toMatch(/parêntese/i)
    expect(problem?.position).toBeGreaterThan(0)
  })

  it('explains the comma used as a separator', () => {
    expect(checkFormula('=SOMA(A1,B1)')?.message).toMatch(/ponto e vírgula/i)
  })
})

describe('names', () => {
  it('rejects a function that does not exist', () => {
    expect(checkFormula('=SOMATÓRIO(A1)')?.message).toMatch(/não conheço a função SOMATÓRIO/i)
  })

  it('accepts the English name', () => {
    expect(checkFormula('=SUM(A1)')).toBeNull()
  })

  it('looks through the whole formula, not just the root', () => {
    expect(checkFormula('=1+INVENTADA(2)')?.message).toMatch(/INVENTADA/)
  })
})

describe('argument count', () => {
  it('complains about an extra argument', () => {
    // Otherwise the user would only see #VALOR! and look for the mistake in the wrong place.
    expect(checkFormula('=ABS(1;2)')?.message).toMatch(/no máximo 1 argumento/i)
  })

  it('complains about a missing argument', () => {
    expect(checkFormula('=ARRED(1)')?.message).toMatch(/precisa de 2 argumentos/i)
  })

  it('says "none" when no argument was given', () => {
    expect(checkFormula('=SOMA()')?.message).toMatch(/recebeu nenhum/i)
  })

  it('also checks the functions that receive their arguments unevaluated', () => {
    expect(checkFormula('=SE(A1>0)')?.message).toMatch(/SE precisa de 2 argumentos/i)
    expect(checkFormula('=SEERRO(A1)')?.message).toMatch(/SEERRO precisa de 2 argumentos/i)
  })

  it('leaves variadic functions alone', () => {
    expect(checkFormula('=CONCATENAR("a";"b";"c";"d")')).toBeNull()
  })
})

describe('in English', () => {
  it.each([
    ['=SOMA(A1,B1)', 'Use semicolons to separate arguments: SUM(A1;B1).'],
    ['=SOMATÓRIO(A1)', 'Unknown function SOMATÓRIO.'],
    ['=ABS(1;2)', 'ABS takes at most 1 argument, and got 2.'],
    ['=ARRED(1)', 'ARRED needs 2 arguments, and got 1 argument.'],
    ['=SUM()', 'SUM needs 1 argument, and got none.'],
  ])('%s explains the problem in English', (formula, message) => {
    expect(checkFormula(formula, Language.English)?.message).toBe(message)
  })
})
