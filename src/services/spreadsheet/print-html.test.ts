import { describe, expect, it } from 'vitest'
import { SHEET_PRINT_CSS, buildSheetHtml, usedBounds } from './print-html.js'
import { createSheet, setCell, type Sheet } from './model.js'
import { CellFormat } from './model.js'

function sheetWith(cells: Array<[number, number, Parameters<typeof setCell>[3]]>): Sheet {
  return cells.reduce((sheet, [row, column, cell]) => setCell(sheet, row, column, cell), createSheet('Plan1'))
}

describe('usedBounds', () => {
  it('para no último dado, não no fim da grade', () => {
    // Uma planilha nova tem mil linhas e nenhum dado. Imprimir a grade inteira
    // gastaria dezenas de páginas em branco, e o usuário descobriria isso na
    // bandeja da impressora.
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

    // O mesmo `formatCell` da grade: a célula não pode sair diferente no papel.
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
    // O Excel joga número à direita por conta própria; a grade daqui não faz
    // isso. Imprimir diferente do que está na tela quebra a única promessa que
    // a impressão tem.
    const sheet = sheetWith([[0, 0, { value: 42 }]])

    expect(buildSheetHtml(sheet)).not.toContain('text-align')
  })

  it('as linhas congeladas viram cabeçalho da tabela', () => {
    // O navegador repete o `<thead>` no topo de cada página impressa. É o que
    // quem congelou a linha na tela esperava que acontecesse no papel.
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
    // O texto vem de um documento, que é dado não confiável.
    const sheet = sheetWith([[0, 0, { value: '<script>alert(1)</script>' }]])

    expect(buildSheetHtml(sheet)).not.toContain('<script>')
    expect(buildSheetHtml(sheet)).toContain('&lt;script&gt;')
  })

  it('recusa cor que não seja #rrggbb', () => {
    // Sem a trava, uma "cor" com ponto e vírgula sairia do atributo `style` e
    // viraria outra declaração.
    const sheet = sheetWith([[0, 0, { value: 'x', style: { color: 'red;background:url(http://x)' } }]])
    const html = buildSheetHtml(sheet)

    expect(html).not.toContain('url(')
    expect(html).toContain('color:inherit')
  })

  it('converte a largura das colunas em proporção', () => {
    // Em pixels, uma planilha mais larga que a página sai com a última coluna
    // cortada na margem — e o usuário só descobre depois de imprimir.
    const sheet: Sheet = { ...sheetWith([[0, 1, { value: 'x' }]]), columnWidths: { 0: 288, 1: 96 } }
    const html = buildSheetHtml(sheet)

    expect(html).toContain('width:75.000%')
    expect(html).toContain('width:25.000%')
    expect(html).not.toContain('px')
  })

  it('encolhe a fonte conforme a planilha alarga', () => {
    // Doze colunas numa A4 em retrato dão cerca de cinquenta pixels cada; em
    // corpo 11 a quebra começa a partir números no meio.
    const estreita = sheetWith([[0, 3, { value: 'x' }]])
    const larga = sheetWith([[0, 14, { value: 'x' }]])

    expect(buildSheetHtml(estreita)).toContain('font-size:11pt')
    expect(buildSheetHtml(larga)).toContain('font-size:8pt')
  })

  it('não corta o texto que não cabe na coluna', () => {
    // Na tela dá para alargar a coluna; no papel, não. Conteúdo escondido no
    // papel é perda silenciosa.
    expect(SHEET_PRINT_CSS).not.toContain('text-overflow')
    expect(SHEET_PRINT_CSS).toContain('word-break')
  })
})
