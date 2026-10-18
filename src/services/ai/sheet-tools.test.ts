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

describe('ferramentas da planilha', () => {
  it('oferecem listar, ler, buscar e propor escrita', () => {
    expect(SHEET_TOOLS.map((tool) => tool.name)).toEqual([
      'list_sheets',
      'read_cells',
      'find_cells',
      'write_cells',
    ])
  })

  it('listam as abas, onde cada uma tem conteúdo e a seleção da ativa', () => {
    expect(call(SheetToolName.ListSheets, {}).content).toBe(
      [
        '1. Vendas (active): content in A1:B2',
        '2. Dados 2026: content in C3',
        '3. Vazia: empty',
        'Selected on the active sheet: A2:B2',
      ].join('\n'),
    )
  })

  it('leem só as células com conteúdo, com a fórmula e o valor, em qualquer aba', () => {
    expect(call(SheetToolName.ReadCells, { range: 'B2:A1' }).content).toBe(
      ['Sheet: Vendas', 'A1: Produto', 'B1: Total', 'A2: Café', 'B2: =2*5 → 10'].join('\n'),
    )
    expect(call(SheetToolName.ReadCells, { range: 'C3', sheet: 'dados 2026' }).content).toContain(
      'C3: café moído',
    )
    expect(call(SheetToolName.ReadCells, { range: 'D1:D9' }).content).toBe('D1:D9 on Vendas is empty.')
  })

  it('recusam intervalo mal escrito, fora da planilha ou aba que não existe', () => {
    expect(call(SheetToolName.ReadCells, { range: 'A1-B2' })).toMatchObject({ isError: true })
    expect(call(SheetToolName.ReadCells, { range: 'A1:A5000' }).content).toContain('outside the sheet Vendas')
    expect(call(SheetToolName.ReadCells, { range: 'A1', sheet: 'Nada' }).content).toBe(
      'There is no sheet named "Nada". The sheets are: Vendas, Dados 2026, Vazia.',
    )
  })

  it('dizem de onde continuar quando a leitura passa do limite', () => {
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

  it('param antes do limite de tamanho quando as células são longas', () => {
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

  it('buscam em todas as abas, ignorando maiúsculas, ou só na pedida', () => {
    expect(call(SheetToolName.FindCells, { query: 'CAFÉ' }).content).toBe(
      ['Vendas!A2: Café', "'Dados 2026'!C3: café moído"].join('\n'),
    )
    expect(call(SheetToolName.FindCells, { query: '2*5', sheet: 'Vendas' }).content).toBe(
      'Vendas!B2: =2*5 → 10',
    )
    expect(call(SheetToolName.FindCells, { query: 'chá' }).content).toBe('No cell contains that text.')
  })

  it('propõem a escrita sem mexer na planilha, na aba pedida', () => {
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

  it('não propõem fórmula que a planilha recusaria nem escrita que passa da borda', () => {
    const invalid = call(SheetToolName.WriteCells, { start: 'A1', rows: [['ok', '=NADA(1)']] })
    expect(invalid).toMatchObject({ isError: true })
    expect(invalid.content).toMatch(/^B1: /)
    expect(call(SheetToolName.WriteCells, { start: 'Z1', rows: [['a', 'b']] }).content).toContain('outside')
    expect(call(SheetToolName.WriteCells, { start: 'A1:B2', rows: [['a']] }).proposal).toBeUndefined()
  })
})
