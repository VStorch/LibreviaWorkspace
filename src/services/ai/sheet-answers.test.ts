import { describe, expect, it } from 'vitest'
import { writeText } from '@services/spreadsheet/edit.js'
import { createSheet, getCell } from '@services/spreadsheet/model.js'
import { applyFill, previewFormula, readFillAnswer, readFormulaAnswer } from './sheet-answers.js'

describe('readFormulaAnswer', () => {
  it('aceita a fórmula sozinha, dentro de cercas de código ou no meio de uma frase', () => {
    const expected = { kind: 'formula', formula: '=SOMA(B2:B3)' }
    expect(readFormulaAnswer('=SOMA(B2:B3)')).toEqual(expected)
    expect(readFormulaAnswer('```excel\n=SOMA(B2:B3)\n```')).toEqual(expected)
    expect(readFormulaAnswer('Use `=SOMA(B2:B3)` na célula.')).toEqual(expected)
    expect(readFormulaAnswer('Aqui está:\n=SOMA(B2:B3).')).toEqual(expected)
  })

  it('um "=" da frase antes da fórmula não a esconde: vale o primeiro trecho que é fórmula válida', () => {
    expect(readFormulaAnswer('Se A1 = 3, use =SOMA(B2:B3)')).toEqual({
      kind: 'formula',
      formula: '=SOMA(B2:B3)',
    })
    expect(readFormulaAnswer('Se A1 = 3, use =XLOOKUP(A1;B1:B9)').kind).toBe('invalid')
  })

  it('sem "=" a resposta não traz fórmula', () => {
    expect(readFormulaAnswer('SOMA(B2:B3)')).toEqual({ kind: 'missing' })
    expect(readFormulaAnswer('')).toEqual({ kind: 'missing' })
  })

  it('função que o Librevia não conhece ou escrita errada vem com o problema', () => {
    const unknown = readFormulaAnswer('=XLOOKUP(A1;B1:B9;C1:C9)')
    expect(unknown).toMatchObject({ kind: 'invalid', formula: '=XLOOKUP(A1;B1:B9;C1:C9)' })
    expect(unknown.kind === 'invalid' && unknown.message).toContain('XLOOKUP')
    expect(readFormulaAnswer('=SOMA(A1,B1)').kind).toBe('invalid')
  })
})

describe('readFillAnswer', () => {
  it('lê um valor por linha, também dentro de cercas', () => {
    expect(readFillAnswer('["a", 2, true, null]', 4)).toEqual(['a', 2, true, null])
    expect(readFillAnswer('```json\n["a","b"]\n```', 2)).toEqual(['a', 'b'])
  })

  it('recusa tamanho errado, tipos errados e texto que não é JSON', () => {
    expect(readFillAnswer('["a"]', 2)).toBeNull()
    expect(readFillAnswer('[["a"], {"b": 1}]', 2)).toBeNull()
    expect(readFillAnswer('a, b', 2)).toBeNull()
    expect(readFillAnswer('[1, 2', 2)).toBeNull()
  })
})

describe('applyFill', () => {
  it('grava valores, e texto que começa com "=" entra como texto, não como fórmula', () => {
    const target = { fromRow: 1, toRow: 3, fromColumn: 2, toColumn: 2 }
    const sheet = applyFill(createSheet('S'), target, ['=1+1', 7, null])
    expect(getCell(sheet, 1, 2)).toEqual({ value: '=1+1' })
    expect(getCell(sheet, 2, 2)).toEqual({ value: 7 })
    expect(getCell(sheet, 3, 2)).toBeUndefined()
  })
})

describe('previewFormula', () => {
  it('calcula numa cópia, sem mudar a pasta aberta', () => {
    const sheet = writeText(writeText(createSheet('S'), 0, 0, '2'), 1, 0, '3')
    const workbook = { sheets: [sheet], activeSheet: 0 }
    expect(previewFormula(workbook, 0, { row: 2, column: 0 }, '=SOMA(A1:A2)')?.value).toBe(5)
    expect(getCell(workbook.sheets[0]!, 2, 0)).toBeUndefined()
  })
})
