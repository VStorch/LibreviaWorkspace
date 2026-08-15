import { describe, expect, it } from 'vitest'
import { adjustForColumns, adjustForRows, translateFormula } from './adjust.js'

const own = { sheet: 'Plan1', own: true }
const other = { sheet: 'Dados', own: false }

describe('copiar a fórmula', () => {
  it('desloca as referências relativas', () => {
    expect(translateFormula('=A1+B2', 1, 0)).toBe('=A2+B3')
  })

  it('não mexe nas absolutas — é para isso que o $ existe', () => {
    expect(translateFormula('=$A$1+A1', 2, 0)).toBe('=$A$1+A3')
  })

  it('trava só o eixo marcado', () => {
    expect(translateFormula('=A$1+$A1', 1, 1)).toBe('=B$1+$A2')
  })

  it('sair da planilha vira #REF!', () => {
    // Copiar =A1 para a coluna A não tem para onde apontar.
    expect(translateFormula('=A1', 0, -1)).toBe('=#REF!')
  })

  it('leva o intervalo inteiro junto', () => {
    expect(translateFormula('=SOMA(A1:A3)', 1, 0)).toBe('=SOMA(A2:A4)')
  })

  it('preserva espaços, maiúsculas e o resto do texto', () => {
    // A reescrita é sobre os símbolos, não sobre a árvore: a fórmula volta como
    // o usuário digitou, menos as referências que precisavam mudar.
    expect(translateFormula('= soma( a1 ; "a1" )', 1, 0)).toBe('= soma( A2 ; "a1" )')
  })
})

describe('inserir linha', () => {
  it('empurra a referência que estava embaixo', () => {
    expect(adjustForRows('=A5', 2, 1, own)).toBe('=A6')
  })

  it('não mexe no que está acima', () => {
    expect(adjustForRows('=A1', 2, 1, own)).toBe('=A1')
  })

  it('mexe também nas absolutas, porque a célula andou de verdade', () => {
    // Aqui o $ não protege: quem estava em $A$5 agora está em $A$6.
    expect(adjustForRows('=$A$5', 2, 1, own)).toBe('=$A$6')
  })

  it('estica o intervalo quando a linha entra no meio dele', () => {
    expect(adjustForRows('=SOMA(A1:A5)', 2, 1, own)).toBe('=SOMA(A1:A6)')
  })

  it('empurra o intervalo inteiro quando entra antes dele', () => {
    expect(adjustForRows('=SOMA(A3:A5)', 0, 1, own)).toBe('=SOMA(A4:A6)')
  })
})

describe('excluir linha', () => {
  it('puxa a referência de baixo', () => {
    expect(adjustForRows('=A5', 0, -1, own)).toBe('=A4')
  })

  it('a referência para a linha excluída vira #REF!', () => {
    expect(adjustForRows('=A3', 2, -1, own)).toBe('=#REF!')
  })

  it('encolhe o intervalo em vez de destruí-lo', () => {
    // A parte mais fácil de errar: tratar as pontas separadamente faria a
    // primeira virar #REF! e mataria a fórmula por uma exclusão que o Excel
    // absorve sem reclamar.
    expect(adjustForRows('=SOMA(A1:A5)', 0, -3, own)).toBe('=SOMA(A1:A2)')
  })

  it('encolhe pelo fim também', () => {
    expect(adjustForRows('=SOMA(A1:A5)', 3, -2, own)).toBe('=SOMA(A1:A3)')
  })

  it('só vira #REF! quando o intervalo inteiro some', () => {
    expect(adjustForRows('=SOMA(A1:A5)', 0, -5, own)).toBe('=SOMA(#REF!:#REF!)')
  })
})

describe('colunas', () => {
  it('inserir empurra para a direita', () => {
    expect(adjustForColumns('=C1', 1, 1, own)).toBe('=D1')
  })

  it('excluir puxa para a esquerda', () => {
    expect(adjustForColumns('=C1', 0, -1, own)).toBe('=B1')
  })

  it('encolhe o intervalo', () => {
    expect(adjustForColumns('=SOMA(A1:E1)', 0, -2, own)).toBe('=SOMA(A1:C1)')
  })
})

describe('entre planilhas', () => {
  it('a fórmula da própria planilha ajusta a referência sem nome', () => {
    expect(adjustForRows('=A5', 0, 1, own)).toBe('=A6')
  })

  it('a fórmula de outra planilha não ajusta a referência sem nome', () => {
    // =A1 na aba "Resumo" aponta para "Resumo", e não para "Dados".
    expect(adjustForRows('=A5', 0, 1, other)).toBe('=A5')
  })

  it('mas ajusta a que nomeia a planilha mexida', () => {
    expect(adjustForRows('=Dados!A5', 0, 1, other)).toBe('=Dados!A6')
  })

  it('não confunde outra planilha qualquer', () => {
    expect(adjustForRows('=Outra!A5', 0, 1, other)).toBe('=Outra!A5')
  })

  it('mantém o nome só na primeira ponta do intervalo', () => {
    expect(adjustForRows('=SOMA(Dados!A1:A5)', 0, 1, other)).toBe('=SOMA(Dados!A2:A6)')
  })

  it('não diferencia maiúsculas no nome da aba', () => {
    expect(adjustForRows('=dados!A5', 0, 1, other)).toBe('=dados!A6')
  })
})

describe('o que não deve ser tocado', () => {
  it('texto que parece referência', () => {
    expect(adjustForRows('=SE(A1="B2";"C3";A5)', 0, 1, own)).toBe('=SE(A2="B2";"C3";A6)')
  })

  it('nome de função com letra e número', () => {
    expect(translateFormula('=HOJE()+A1', 1, 0)).toBe('=HOJE()+A2')
  })

  it('fórmula que nem chega a ser lida volta intacta', () => {
    expect(adjustForRows('=SOMA("', 0, 1, own)).toBe('=SOMA("')
  })

  it('fórmula sem o sinal de igual também funciona', () => {
    expect(adjustForRows('A5', 0, 1, own)).toBe('A6')
  })
})
