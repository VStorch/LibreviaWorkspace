import { describe, expect, it } from 'vitest'
import { parseClipboardText, pasteText, rangeAsText } from './clipboard.js'
import { createSheet, getCell, setCell, type Sheet } from './model.js'

function filled(): Sheet {
  let sheet = createSheet('S')
  sheet = setCell(sheet, 0, 0, { value: 'a' })
  sheet = setCell(sheet, 0, 2, { formula: '=1+1', value: 2 })
  sheet = setCell(sheet, 1, 1, { value: 'linha\ncom "aspas"' })
  return sheet
}

describe('rangeAsText', () => {
  it('copia o que a célula mostra, com tabulação entre colunas e uma linha por linha', () => {
    expect(rangeAsText(filled(), { fromRow: 0, fromColumn: 0, toRow: 0, toColumn: 2 })).toBe('a\t\t2')
  })

  it('põe entre aspas a célula com quebra de linha ou aspas, como o Excel', () => {
    const range = { fromRow: 1, fromColumn: 0, toRow: 1, toColumn: 1 }
    expect(rangeAsText(filled(), range)).toBe('\t"linha\ncom ""aspas"""')
  })

  it('aceita o intervalo de trás para a frente', () => {
    expect(rangeAsText(filled(), { fromRow: 0, fromColumn: 2, toRow: 0, toColumn: 0 })).toBe('a\t\t2')
  })
})

describe('parseClipboardText', () => {
  it('separa linhas e colunas e ignora a quebra de linha final', () => {
    expect(parseClipboardText('a\tb\nc\td\n')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
    expect(parseClipboardText('a\r\nb')).toEqual([['a'], ['b']])
  })

  it('mantém células vazias e lê as aspas de volta', () => {
    expect(parseClipboardText('\t"linha\ncom ""aspas"""\t')).toEqual([['', 'linha\ncom "aspas"', '']])
  })

  it('é o inverso de rangeAsText', () => {
    const range = { fromRow: 0, fromColumn: 0, toRow: 1, toColumn: 2 }
    const text = rangeAsText(filled(), range)
    expect(parseClipboardText(text)).toEqual([
      ['a', '', '2'],
      ['', 'linha\ncom "aspas"', ''],
    ])
  })
})

describe('pasteText', () => {
  it('escreve a partir da célula, como se digitasse, e descarta o que passa da planilha', () => {
    const sheet = createSheet('S')
    const corner = sheet.columnCount - 1
    const pasted = pasteText(sheet, 0, corner - 1, [
      ['=1+2', 'x', 'fora'],
      ['3', 'y'],
    ])
    expect(getCell(pasted, 0, corner - 1)?.formula).toBe('=1+2')
    expect(getCell(pasted, 0, corner)?.value).toBe('x')
    expect(getCell(pasted, 1, corner - 1)?.value).toBe(3)
    expect(Object.keys(pasted.cells)).toHaveLength(4)
  })
})
