import { describe, expect, it } from 'vitest'
import { createSheet } from './model.js'
import { workbookStatistics } from './statistics.js'

describe('workbookStatistics', () => {
  it('conta as abas, as células preenchidas e as fórmulas de todas as abas', () => {
    const sales = {
      ...createSheet('Vendas'),
      cells: { A1: { value: 'Cabo' }, B1: { value: 3 }, C1: { formula: '=B1*2', value: 6 } },
    }
    const summary = { ...createSheet('Resumo'), cells: { A1: { formula: '=Vendas!C1', value: 6 } } }

    expect(workbookStatistics({ sheets: [sales, summary], activeSheet: 0 })).toEqual({
      sheets: 2,
      filledCells: 4,
      formulas: 2,
    })
  })

  it('não conta a célula que só tem formatação ou texto vazio', () => {
    const sheet = { ...createSheet('A'), cells: { A1: { style: { bold: true } }, A2: { value: '' } } }

    expect(workbookStatistics({ sheets: [sheet], activeSheet: 0 }).filledCells).toBe(0)
  })
})
