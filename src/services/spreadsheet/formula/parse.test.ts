import { describe, expect, it } from 'vitest'
import { ParseError } from './errors.js'
import { formatReference } from './references.js'
import { parseFormula, tryParseFormula } from './parse.js'
import { walk, type Node } from './ast.js'

/** A árvore como texto, para o teste falar da forma e não da estrutura. */
function show(node: Node): string {
  switch (node.kind) {
    case 'number':
      return String(node.value)
    case 'text':
      return `"${node.value}"`
    case 'boolean':
      return node.value ? 'VERDADEIRO' : 'FALSO'
    case 'error':
      return node.value
    case 'reference':
      return formatReference(node.ref)
    case 'range':
      return `${formatReference(node.from)}:${formatReference(node.to)}`
    case 'unary':
      return `(${node.operator}${show(node.operand)})`
    case 'percent':
      return `(${show(node.operand)}%)`
    case 'binary':
      return `(${show(node.left)} ${node.operator} ${show(node.right)})`
    case 'call':
      return `${node.name}(${node.args.map(show).join('; ')})`
  }
}

const tree = (formula: string): string => show(parseFormula(formula))

describe('precedência', () => {
  it.each([
    ['=1+2*3', '(1 + (2 * 3))'],
    ['=(1+2)*3', '((1 + 2) * 3)'],
    ['=1+2-3', '((1 + 2) - 3)'],
    ['=1&2&3', '((1 & 2) & 3)'],
    ['=1<2+3', '(1 < (2 + 3))'],
    ['=1&2=3', '((1 & 2) = 3)'],
  ])('%s vira %s', (formula, expected) => {
    expect(tree(formula)).toBe(expected)
  })

  it('potência associa à direita', () => {
    // 2^(3^2) = 512, e não (2^3)^2 = 64.
    expect(tree('=2^3^2')).toBe('(2 ^ (3 ^ 2))')
  })

  it('menos unário liga mais forte que a potência, como no Excel', () => {
    // =-2^2 dá 4 no Excel, e não -4 como na matemática. O resultado precisa
    // bater com a planilha ao lado, não com o livro.
    expect(tree('=-2^2')).toBe('((-2) ^ 2)')
  })

  it('aceita expoente negativo', () => {
    expect(tree('=2^-1')).toBe('(2 ^ (-1))')
  })
})

describe('números', () => {
  it('lê vírgula como decimal', () => {
    expect(tree('=1,5')).toBe('1.5')
  })

  it('aceita ponto decimal também, que é o que vem colado de fora', () => {
    expect(tree('=1.5')).toBe('1.5')
  })

  it('lê notação científica', () => {
    expect(tree('=2e-3')).toBe('0.002')
  })

  it('trata % como sufixo', () => {
    expect(tree('=50%')).toBe('(50%)')
    expect(tree('=A1*10%')).toBe('(A1 * (10%))')
  })
})

describe('texto', () => {
  it('lê aspas duplicadas como uma aspa', () => {
    expect(tree('="diz ""oi"""')).toBe('"diz "oi""')
  })

  it('reclama de aspas não fechadas', () => {
    expect(() => parseFormula('="aberto')).toThrow(ParseError)
  })

  it('não confunde ponto e vírgula dentro do texto com separador', () => {
    expect(tree('=SOMA("a;b")')).toBe('SOMA("a;b")')
  })
})

describe('referências', () => {
  it.each([
    ['=A1', 'A1'],
    ['=$A$1', '$A$1'],
    ['=A$1', 'A$1'],
    ['=$A1', '$A1'],
    ['=Plan1!A1', 'Plan1!A1'],
    ["='Vendas 2026'!A1", "'Vendas 2026'!A1"],
    ['=AA100', 'AA100'],
  ])('%s vira %s', (formula, expected) => {
    expect(tree(formula)).toBe(expected)
  })

  it('normaliza o intervalo escrito ao contrário', () => {
    // B4:A1 e A1:B4 são o mesmo retângulo; quem lê a árvore não deveria
    // precisar saber em que ordem foi digitado.
    expect(tree('=SOMA(B4:A1)')).toBe('SOMA(A1:B4)')
  })

  it('leva o $ para o canto certo ao normalizar', () => {
    expect(tree('=SOMA($B$4:A1)')).toBe('SOMA(A1:$B$4)')
  })

  it('mantém o nome da planilha no intervalo', () => {
    expect(tree('=SOMA(Plan2!A1:B2)')).toBe('SOMA(Plan2!A1:B2)')
  })
})

describe('funções', () => {
  it('lê chamada sem argumentos', () => {
    expect(tree('=HOJE()')).toBe('HOJE()')
  })

  it('separa argumentos por ponto e vírgula', () => {
    expect(tree('=SE(A1>0;"sim";"não")')).toBe('SE((A1 > 0); "sim"; "não")')
  })

  it('aceita nome com ponto', () => {
    expect(tree('=CONT.SE(A1:A9;">0")')).toBe('CONT.SE(A1:A9; ">0")')
  })

  it('aninha chamadas', () => {
    expect(tree('=SOMA(MÁXIMO(A1;B1);2)')).toBe('SOMA(MÁXIMO(A1; B1); 2)')
  })

  it('maiúsculas e minúsculas dão no mesmo', () => {
    expect(tree('=soma(a1:b2)')).toBe('SOMA(A1:B2)')
  })
})

describe('erros de escrita', () => {
  it.each([
    ['=SOMA(A1', 'parêntese'],
    ['=1+', 'terminou'],
    ['=A1 B1', 'Sobrou'],
    ['=', 'vazia'],
  ])('%s reclama', (formula, trecho) => {
    expect(() => parseFormula(formula)).toThrow(new RegExp(trecho, 'i'))
  })

  it('aponta a vírgula usada como separador', () => {
    // O erro mais provável de quem vem do Excel em inglês. Dizer só "não
    // entendi" mandaria o usuário procurar no lugar errado.
    expect(() => parseFormula('=SOMA(A1,B1)')).toThrow(/ponto e vírgula/i)
  })

  it('diz onde está o problema', () => {
    try {
      parseFormula('=SOMA(A1;)')
      expect.unreachable('deveria ter reclamado')
    } catch (error) {
      expect(error).toBeInstanceOf(ParseError)
      expect((error as ParseError).position).toBeGreaterThan(0)
    }
  })

  it('a versão silenciosa devolve nulo em vez de lançar', () => {
    expect(tryParseFormula('=SOMA(')).toBeNull()
    expect(tryParseFormula('=1')).not.toBeNull()
  })
})

describe('erros literais', () => {
  it('aceita #N/D escrito na fórmula', () => {
    expect(tree('=SEERRO(A1;#N/D)')).toBe('SEERRO(A1; #N/D)')
  })
})

describe('percorrer a árvore', () => {
  it('visita todos os nós', () => {
    const refs = [...walk(parseFormula('=SOMA(A1:B2;C3)*D4'))].filter(
      (node) => node.kind === 'reference' || node.kind === 'range',
    )

    expect(refs.map(show)).toEqual(['A1:B2', 'C3', 'D4'])
  })
})
