import { describe, expect, it } from 'vitest'
import { SHEET_PRINT_CSS, buildSheetHtml, usedBounds } from './print-html.js'
import { createSheet, setCell, type Sheet } from './model.js'
import { CellFormat } from './model.js'

function sheetWith(cells: Array<[number, number, Parameters<typeof setCell>[3]]>): Sheet {
  return cells.reduce((sheet, [row, column, cell]) => setCell(sheet, row, column, cell), createSheet('Plan1'))
}

describe('usedBounds', () => {
  it('para no último dado, não no fim da grade', () => {
    // A new sheet has a thousand rows and no data. Printing the whole grid would waste dozens of
    // blank pages, and the user would find out at the printer tray.
    const sheet = sheetWith([
      [0, 0, { value: 'Produto' }],
      [2, 3, { value: 10 }],
    ])

    expect(usedBounds(sheet)).toEqual({ rows: 3, columns: 4 })
  })

  it('planilha vazia não tem área de impressão', () => {
    expect(usedBounds(createSheet('Plan1'))).toEqual({ rows: 0, columns: 0 })
  })
})

describe('buildSheetHtml', () => {
  it('diz que a aba está vazia em vez de imprimir uma tabela sem linhas', () => {
    expect(buildSheetHtml(createSheet('Plan1'))).toContain('vazia')
  })

  it('formata a célula como na tela', () => {
    const sheet = sheetWith([[0, 0, { value: 12.5, style: { format: CellFormat.Currency } }]])

    // The same `formatCell` as the grid: the cell must not look different on paper.
    expect(buildSheetHtml(sheet)).toContain('R$')
  })

  it('leva negrito, cor e alinhamento para o papel', () => {
    const sheet = sheetWith([
      [0, 0, { value: 'Título', style: { bold: true, color: '#1a5fb4', align: 'center' } }],
    ])
    const html = buildSheetHtml(sheet)

    expect(html).toContain('font-weight:700')
    expect(html).toContain('color:#1a5fb4')
    expect(html).toContain('text-align:center')
  })

  it('não inventa alinhamento que a tela não faz', () => {
    // Excel right-aligns numbers on its own; this grid does not. Printing differently from the
    // screen breaks the only promise printing has.
    const sheet = sheetWith([[0, 0, { value: 42 }]])

    expect(buildSheetHtml(sheet)).not.toContain('text-align')
  })

  it('as linhas congeladas viram cabeçalho da tabela', () => {
    // The browser repeats `<thead>` at the top of each printed page. That is what someone who froze
    // the row on screen expects on paper.
    const sheet: Sheet = {
      ...sheetWith([
        [0, 0, { value: 'Produto' }],
        [1, 0, { value: 'Cabo' }],
      ]),
      frozenRows: 1,
    }
    const html = buildSheetHtml(sheet)

    expect(html).toMatch(/<thead>.*Produto.*<\/thead>/s)
    expect(html).toMatch(/<tbody>.*Cabo.*<\/tbody>/s)
  })

  it('sem congelamento não inventa cabeçalho', () => {
    const sheet = sheetWith([[0, 0, { value: 'Produto' }]])

    expect(buildSheetHtml(sheet)).not.toContain('<thead>')
  })

  it('escapa o conteúdo da célula', () => {
    // The text comes from a document, which is untrusted data.
    const sheet = sheetWith([[0, 0, { value: '<script>alert(1)</script>' }]])

    expect(buildSheetHtml(sheet)).not.toContain('<script>')
    expect(buildSheetHtml(sheet)).toContain('&lt;script&gt;')
  })

  it('recusa cor que não seja #rrggbb', () => {
    // Without the guard, a "color" with a semicolon would leave the `style` attribute and become
    // another declaration.
    const sheet = sheetWith([[0, 0, { value: 'x', style: { color: 'red;background:url(http://x)' } }]])
    const html = buildSheetHtml(sheet)

    expect(html).not.toContain('url(')
    expect(html).toContain('color:inherit')
  })

  it('converte a largura das colunas em proporção', () => {
    // In pixels, a sheet wider than the page comes out with the last column clipped at the margin,
    // and the user only finds out after printing.
    const sheet: Sheet = { ...sheetWith([[0, 1, { value: 'x' }]]), columnWidths: { 0: 288, 1: 96 } }
    const html = buildSheetHtml(sheet)

    expect(html).toContain('width:75.000%')
    expect(html).toContain('width:25.000%')
    expect(html).not.toContain('px')
  })

  it('encolhe a fonte conforme a planilha alarga', () => {
    // Twelve columns on portrait A4 give about fifty pixels each; at 11 pt numbers start breaking
    // in the middle.
    const estreita = sheetWith([[0, 3, { value: 'x' }]])
    const larga = sheetWith([[0, 14, { value: 'x' }]])

    expect(buildSheetHtml(estreita)).toContain('font-size:11pt')
    expect(buildSheetHtml(larga)).toContain('font-size:8pt')
  })

  it('não corta o texto que não cabe na coluna', () => {
    // On screen the column can be widened; on paper it cannot. Content hidden on paper is silent
    // loss.
    expect(SHEET_PRINT_CSS).not.toContain('text-overflow')
    expect(SHEET_PRINT_CSS).toContain('word-break')
  })
})
