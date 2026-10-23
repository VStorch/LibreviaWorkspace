import { describe, expect, it } from 'vitest'
import { Language } from '@shared/i18n/index.js'
import { ParseError } from './errors.js'
import { formatReference } from './references.js'
import { parseFormula, tryParseFormula } from './parse.js'
import { walk, type Node } from './ast.js'

/** The tree as text, so the test speaks about shape and not structure. */
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

describe('precedence', () => {
  it.each([
    ['=1+2*3', '(1 + (2 * 3))'],
    ['=(1+2)*3', '((1 + 2) * 3)'],
    ['=1+2-3', '((1 + 2) - 3)'],
    ['=1&2&3', '((1 & 2) & 3)'],
    ['=1<2+3', '(1 < (2 + 3))'],
    ['=1&2=3', '((1 & 2) = 3)'],
  ])('%s becomes %s', (formula, expected) => {
    expect(tree(formula)).toBe(expected)
  })

  it('exponentiation is right-associative', () => {
    // 2^(3^2) = 512, not (2^3)^2 = 64.
    expect(tree('=2^3^2')).toBe('(2 ^ (3 ^ 2))')
  })

  it('unary minus binds tighter than exponentiation, as in Excel', () => {
    // =-2^2 is 4 in Excel, not -4 as in math. The result must match the spreadsheet next to it, not
    // the textbook.
    expect(tree('=-2^2')).toBe('((-2) ^ 2)')
  })

  it('accepts a negative exponent', () => {
    expect(tree('=2^-1')).toBe('(2 ^ (-1))')
  })
})

describe('numbers', () => {
  it('reads a comma as decimal', () => {
    expect(tree('=1,5')).toBe('1.5')
  })

  it('accepts a decimal point too, which is what comes pasted from outside', () => {
    expect(tree('=1.5')).toBe('1.5')
  })

  it('reads scientific notation', () => {
    expect(tree('=2e-3')).toBe('0.002')
  })

  it('treats % as a suffix', () => {
    expect(tree('=50%')).toBe('(50%)')
    expect(tree('=A1*10%')).toBe('(A1 * (10%))')
  })
})

describe('text', () => {
  it('reads doubled quotes as one quote', () => {
    expect(tree('="diz ""oi"""')).toBe('"diz "oi""')
  })

  it('complains about unclosed quotes', () => {
    expect(() => parseFormula('="aberto')).toThrow(ParseError)
  })

  it('does not mistake a semicolon inside text for a separator', () => {
    expect(tree('=SOMA("a;b")')).toBe('SOMA("a;b")')
  })
})

describe('references', () => {
  it.each([
    ['=A1', 'A1'],
    ['=$A$1', '$A$1'],
    ['=A$1', 'A$1'],
    ['=$A1', '$A1'],
    ['=Plan1!A1', 'Plan1!A1'],
    ["='Vendas 2026'!A1", "'Vendas 2026'!A1"],
    ['=AA100', 'AA100'],
  ])('%s becomes %s', (formula, expected) => {
    expect(tree(formula)).toBe(expected)
  })

  it('normalizes a range written backwards', () => {
    // B4:A1 and A1:B4 are the same rectangle; whoever reads the tree should not need to know the
    // typing order.
    expect(tree('=SOMA(B4:A1)')).toBe('SOMA(A1:B4)')
  })

  it('moves the $ to the right corner when normalizing', () => {
    expect(tree('=SOMA($B$4:A1)')).toBe('SOMA(A1:$B$4)')
  })

  it('keeps the sheet name in the range', () => {
    expect(tree('=SOMA(Plan2!A1:B2)')).toBe('SOMA(Plan2!A1:B2)')
  })
})

describe('functions', () => {
  it('reads a call without arguments', () => {
    expect(tree('=HOJE()')).toBe('HOJE()')
  })

  it('separates arguments by semicolons', () => {
    expect(tree('=SE(A1>0;"sim";"não")')).toBe('SE((A1 > 0); "sim"; "não")')
  })

  it('accepts a name with a dot', () => {
    expect(tree('=CONT.SE(A1:A9;">0")')).toBe('CONT.SE(A1:A9; ">0")')
  })

  it('nests calls', () => {
    expect(tree('=SOMA(MÁXIMO(A1;B1);2)')).toBe('SOMA(MÁXIMO(A1; B1); 2)')
  })

  it('upper and lower case are equivalent', () => {
    expect(tree('=soma(a1:b2)')).toBe('SOMA(A1:B2)')
  })
})

function problemOf(formula: string): string | undefined {
  try {
    parseFormula(formula)
    return undefined
  } catch (error) {
    if (error instanceof ParseError) return error.key
    throw error
  }
}

describe('syntax errors', () => {
  it.each([
    ['=SOMA(A1', 'spreadsheet.formula.unclosedCall'],
    ['=(1', 'spreadsheet.formula.unclosedParenthesis'],
    ['=1+', 'spreadsheet.formula.endedEarly'],
    ['=A1 B1', 'spreadsheet.formula.trailingToken'],
    ['=', 'spreadsheet.formula.empty'],
    ['="open', 'spreadsheet.formula.unclosedText'],
    ["='Sales!A1", 'spreadsheet.formula.unclosedSheetName'],
    ['=A1:Sheet2!B2', 'spreadsheet.formula.rangeAcrossSheets'],
    ['=#WRONG!', 'spreadsheet.formula.unknownErrorAt'],
  ])('%s complains with %s', (formula, key) => {
    expect(problemOf(formula)).toBe(key)
  })

  it('points out the comma used as a separator', () => {
    // The most likely mistake of someone coming from English Excel. Saying only "did not
    // understand" would send the user to look in the wrong place.
    expect(problemOf('=SOMA(A1,B1)')).toBe('spreadsheet.formula.commaSeparator')
  })

  it('reads the problem in either language', () => {
    try {
      parseFormula('=A1 B1')
      expect.unreachable('should have complained')
    } catch (error) {
      expect((error as ParseError).describe(Language.Portuguese)).toBe('Sobrou "B1" no fim da fórmula.')
      expect((error as ParseError).describe(Language.English)).toBe(
        '"B1" is left over at the end of the formula.',
      )
    }
  })

  it('says where the problem is', () => {
    try {
      parseFormula('=SOMA(A1;)')
      expect.unreachable('should have complained')
    } catch (error) {
      expect(error).toBeInstanceOf(ParseError)
      expect((error as ParseError).position).toBeGreaterThan(0)
    }
  })

  it('the silent version returns null instead of throwing', () => {
    expect(tryParseFormula('=SOMA(')).toBeNull()
    expect(tryParseFormula('=1')).not.toBeNull()
  })
})

describe('literal errors', () => {
  it('accepts #N/D written in the formula', () => {
    expect(tree('=SEERRO(A1;#N/D)')).toBe('SEERRO(A1; #N/D)')
  })
})

describe('walking the tree', () => {
  it('visits every node', () => {
    const refs = [...walk(parseFormula('=SOMA(A1:B2;C3)*D4'))].filter(
      (node) => node.kind === 'reference' || node.kind === 'range',
    )

    expect(refs.map(show)).toEqual(['A1:B2', 'C3', 'D4'])
  })
})
