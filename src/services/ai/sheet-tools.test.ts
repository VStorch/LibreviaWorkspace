import { describe, expect, it } from 'vitest'
import { singleCell } from '@services/spreadsheet/edit.js'
import { createSheet, setCell, type WorkbookModel } from '@services/spreadsheet/model.js'
import { MAX_AI_READ_CELLS, MAX_AI_TOOL_RESULT_LENGTH } from '@shared/limits.js'
import { runSheetTool, SHEET_TOOLS, SheetToolName, type SheetToolContext } from './sheet-tools.js'

function context(): SheetToolContext {
  let sales = createSheet('Vendas')
  sales = setCell(sales, 0, 0, { value: 'Produto' })
  sales = setCell(sales, 0, 1, { value: 'Total' })
  sales = setCell(sales, 1, 0, { value: 'Café' })
  sales = setCell(sales, 1, 1, { formula: '=2*5', value: 10 })
  const data = setCell(createSheet('Dados 2026'), 2, 2, { value: 'café moído' })
  const workbook: WorkbookModel = { sheets: [sales, data, createSheet('Vazia')], activeSheet: 0 }
  return { workbook, selection: { fromRow: 1, fromColumn: 0, toRow: 1, toColumn: 1 } }
}

const call = (name: string, args: unknown, given = context()) =>
  runSheetTool({ id: '1', name, arguments: JSON.stringify(args) }, given)

describe('sheet tools', () => {
  it('offer list, read, search and propose write', () => {
    expect(SHEET_TOOLS.map((tool) => tool.name)).toEqual([
      'list_sheets',
      'read_cells',
      'find_cells',
      'write_cells',
    ])
  })

  it('list the tabs, where each has content and the active selection', () => {
    expect(call(SheetToolName.ListSheets, {}).content).toBe(
      [
        '1. Vendas (active): content in A1:B2',
        '2. Dados 2026: content in C3',
        '3. Vazia: empty',
        'Selected on the active sheet: A2:B2',
      ].join('\n'),
    )
  })

  it('read only cells with content, with formula and value, in any tab', () => {
    expect(call(SheetToolName.ReadCells, { range: 'B2:A1' }).content).toBe(
      ['Sheet: Vendas', 'A1: Produto', 'B1: Total', 'A2: Café', 'B2: =2*5 → 10'].join('\n'),
    )
    expect(call(SheetToolName.ReadCells, { range: 'C3', sheet: 'dados 2026' }).content).toContain(
      'C3: café moído',
    )
    expect(call(SheetToolName.ReadCells, { range: 'D1:D9' }).content).toBe('D1:D9 on Vendas is empty.')
  })

  it('refuse a malformed range, one outside the sheet or a tab that does not exist', () => {
    expect(call(SheetToolName.ReadCells, { range: 'A1-B2' })).toMatchObject({ isError: true })
    expect(call(SheetToolName.ReadCells, { range: 'A1:A5000' }).content).toContain('outside the sheet Vendas')
    expect(call(SheetToolName.ReadCells, { range: 'A1', sheet: 'Nada' }).content).toBe(
      'There is no sheet named "Nada". The sheets are: Vendas, Dados 2026, Vazia.',
    )
  })

  it('say where to continue when the read exceeds the limit', () => {
    let sheet = createSheet('Longa')
    for (let row = 0; row <= MAX_AI_READ_CELLS; row++) sheet = setCell(sheet, row, 0, { value: row })
    const long: SheetToolContext = {
      workbook: { sheets: [sheet], activeSheet: 0 },
      selection: singleCell(0, 0),
    }
    const lines = call(SheetToolName.ReadCells, { range: 'A1:A1000' }, long).content.split('\n')
    expect(lines).toHaveLength(MAX_AI_READ_CELLS + 2)
    expect(lines.at(-1)).toContain(`the next is A${MAX_AI_READ_CELLS + 1}`)
  })

  it('stop before the size limit when cells are long', () => {
    let sheet = createSheet('Textos')
    for (let row = 0; row < MAX_AI_READ_CELLS; row++) {
      sheet = setCell(sheet, row, 0, { formula: `="${'a'.repeat(300)}"`, value: 'b'.repeat(300) })
    }
    const long: SheetToolContext = {
      workbook: { sheets: [sheet], activeSheet: 0 },
      selection: singleCell(0, 0),
    }
    const content = call(SheetToolName.ReadCells, { range: 'A1:A1000' }, long).content
    expect(content.length).toBeLessThanOrEqual(MAX_AI_TOOL_RESULT_LENGTH)
    expect(content.split('\n').at(-1)).toMatch(
      /^Only the first \d+ cells with content are shown; the next is A\d+/,
    )
  })

  it('search all tabs, ignoring case, or only the requested one', () => {
    expect(call(SheetToolName.FindCells, { query: 'CAFÉ' }).content).toBe(
      ['Vendas!A2: Café', "'Dados 2026'!C3: café moído"].join('\n'),
    )
    expect(call(SheetToolName.FindCells, { query: '2*5', sheet: 'Vendas' }).content).toBe(
      'Vendas!B2: =2*5 → 10',
    )
    expect(call(SheetToolName.FindCells, { query: 'chá' }).content).toBe('No cell contains that text.')
  })

  it('propose the write without touching the sheet, in the requested tab', () => {
    const given = context()
    const outcome = call(
      SheetToolName.WriteCells,
      { start: 'c1', rows: [['Preço'], ['=B2/2']], reason: 'metade' },
      given,
    )
    expect(outcome).toMatchObject({
      isError: false,
      proposal: { sheet: 'Vendas', row: 0, column: 2, reason: 'metade' },
    })
    expect(outcome.proposal?.rows).toEqual([['Preço'], ['=B2/2']])
    expect(outcome.content).toContain('C1:C2 on Vendas is waiting')
    expect(given.workbook.sheets[0]!.cells['C1']).toBeUndefined()
    expect(
      call(SheetToolName.WriteCells, { start: 'A1', rows: [['x']], sheet: 'Vazia' }).proposal?.sheet,
    ).toBe('Vazia')
  })

  it('do not propose a formula the sheet would refuse nor a write that goes past the edge', () => {
    const invalid = call(SheetToolName.WriteCells, { start: 'A1', rows: [['ok', '=NADA(1)']] })
    expect(invalid).toMatchObject({ isError: true })
    expect(invalid.content).toBe('B1: Unknown function NADA.')
    expect(call(SheetToolName.WriteCells, { start: 'Z1', rows: [['a', 'b']] }).content).toContain('outside')
    expect(call(SheetToolName.WriteCells, { start: 'A1:B2', rows: [['a']] }).proposal).toBeUndefined()
  })
})
