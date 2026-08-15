import { describe, expect, it } from 'vitest'
import { checkFormula } from './validate.js'

describe('fórmula boa', () => {
  it.each(['=1+1', '=SOMA(A1:A9)', '=SE(A1>0;"sim";"não")', '=HOJE()', '=SOMA(1;2;3;4;5)'])(
    '%s passa',
    (formula) => {
      expect(checkFormula(formula)).toBeNull()
    },
  )
})

describe('problemas de escrita', () => {
  it('aponta a posição', () => {
    const problem = checkFormula('=SOMA(A1')

    expect(problem?.message).toMatch(/parêntese/i)
    expect(problem?.position).toBeGreaterThan(0)
  })

  it('explica a vírgula usada como separador', () => {
    expect(checkFormula('=SOMA(A1,B1)')?.message).toMatch(/ponto e vírgula/i)
  })
})

describe('nomes', () => {
  it('recusa função que não existe', () => {
    expect(checkFormula('=SOMATÓRIO(A1)')?.message).toMatch(/não conheço a função SOMATÓRIO/i)
  })

  it('aceita o nome em inglês', () => {
    expect(checkFormula('=SUM(A1)')).toBeNull()
  })

  it('procura dentro da fórmula inteira, e não só na raiz', () => {
    expect(checkFormula('=1+INVENTADA(2)')?.message).toMatch(/INVENTADA/)
  })
})

describe('número de argumentos', () => {
  it('reclama de argumento a mais', () => {
    // Sem isso o usuário veria só #VALOR! e procuraria o erro no lugar errado.
    expect(checkFormula('=ABS(1;2)')?.message).toMatch(/no máximo 1 argumento/i)
  })

  it('reclama de argumento a menos', () => {
    expect(checkFormula('=ARRED(1)')?.message).toMatch(/precisa de 2 argumentos/i)
  })

  it('diz "nenhum" quando não veio nenhum', () => {
    expect(checkFormula('=SOMA()')?.message).toMatch(/recebeu nenhum/i)
  })

  it('confere também as que recebem os argumentos sem avaliar', () => {
    expect(checkFormula('=SE(A1>0)')?.message).toMatch(/SE precisa de 2 argumentos/i)
    expect(checkFormula('=SEERRO(A1)')?.message).toMatch(/SEERRO precisa de 2 argumentos/i)
  })

  it('deixa as variádicas em paz', () => {
    expect(checkFormula('=CONCATENAR("a";"b";"c";"d")')).toBeNull()
  })
})
