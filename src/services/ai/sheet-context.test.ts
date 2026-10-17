import { describe, expect, it } from 'vitest'
import { writeText } from '@services/spreadsheet/edit.js'
import { createSheet, type Sheet } from '@services/spreadsheet/model.js'
import { describeRowsToFill, describeSheetAround, fitLines } from './sheet-context.js'

function sheetWith(name: string, cells: Record<string, string>): Sheet {
  return Object.entries(cells).reduce((sheet, [ref, text]) => {
    const column = ref.charCodeAt(0) - 65
    const row = Number(ref.slice(1)) - 1
    return writeText(sheet, row, column, text)
  }, createSheet(name))
}

const VENDAS = sheetWith('Vendas', {
  A1: 'Produto',
  B1: 'Preço',
  A2: 'Caneta',
  B2: '3',
  A3: 'Lápis',
  B3: '2',
})

describe('describeSheetAround', () => {
  it('leva o nome da planilha, a célula ativa, os cabeçalhos e as células usadas', () => {
    const text = describeSheetAround(VENDAS, { row: 3, column: 1 }, 10_000)
    expect(text).toBe(
      [
        'Sheet: Vendas',
        'Active cell: B4',
        'Headers:',
        'A1: Produto',
        'B1: Preço',
        'Cells:',
        'B3: 2',
        'B2: 3',
        'A3: Lápis',
        'A2: Caneta',
      ].join('\n'),
    )
  })

  it('cabe no orçamento cortando as células mais distantes', () => {
    const text = describeSheetAround(VENDAS, { row: 3, column: 1 }, 90)
    expect(text.length).toBeLessThanOrEqual(90)
    expect(text).toContain('B3: 2')
    expect(text).not.toContain('A2: Caneta')
  })

  it('mostra a fórmula e o valor, e uma quebra de linha da célula não vira outra linha', () => {
    const sheet = writeText(sheetWith('S', { A1: 'um\ndois' }), 1, 0, '=A1')
    const text = describeSheetAround(sheet, { row: 1, column: 0 }, 10_000)
    expect(text).toContain('A1: um dois')
    expect(text).toContain('A2: =A1 → ')
  })
})

describe('describeRowsToFill', () => {
  it('lista cada linha com as outras colunas e os cabeçalhos acima', () => {
    const text = describeRowsToFill(VENDAS, { fromRow: 1, toRow: 3, fromColumn: 2, toColumn: 2 })
    expect(text).toBe(
      [
        'Sheet: Vendas',
        'Target column: C',
        'Columns: A (Produto) | B (Preço)',
        'Rows (3):',
        'Row 2: A2: Caneta | B2: 3',
        'Row 3: A3: Lápis | B3: 2',
        'Row 4: (empty)',
      ].join('\n'),
    )
  })

  it('sem linha acima do destino, não inventa cabeçalho', () => {
    const text = describeRowsToFill(VENDAS, { fromRow: 0, toRow: 0, fromColumn: 1, toColumn: 1 })
    expect(text).toContain('Target column: B\nColumns: A\n')
  })
})

describe('fitLines', () => {
  it('mantém linhas inteiras e não passa do orçamento', () => {
    expect(fitLines(['abc', 'def'], 100)).toBe('abc\ndef')
    expect(fitLines(['abc', 'def'], 5)).toBe('abc')
  })

  it('avisa que só vão as células mais próximas quando o aviso cabe', () => {
    expect(fitLines(['a', 'b', 'c'.repeat(100)], 60)).toBe(
      'a\nb\nOnly the cells nearest the active one are shown.',
    )
  })
})
