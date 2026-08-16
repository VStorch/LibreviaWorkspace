/**
 * Planilha → HTML de impressão.
 *
 * O documento imprime o HTML que o próprio editor produziu, e por um bom motivo:
 * o que sai no papel é literalmente o que estava na tela. A planilha não pode
 * fazer o mesmo. A grade só desenha as células visíveis — rolar uma planilha de
 * 10 mil linhas mantém cerca de 300 células no DOM — então imprimir o que está
 * lá renderizaria a janela, não a planilha.
 *
 * Por isso aqui o HTML é gerado a partir do **modelo**, com as mesmas funções de
 * formatação que a tela usa (`formatCell`). O que muda é a origem, não as
 * regras: a mesma célula sai igual nos dois lugares.
 */

import { formatCell } from './format.js'
import { cellRef, DEFAULT_COLUMN_WIDTH, type Cell, type Sheet } from './model.js'

/** Até onde imprimir. */
export interface PrintBounds {
  readonly rows: number
  readonly columns: number
}

/**
 * O retângulo que de fato tem conteúdo.
 *
 * Uma planilha nova tem mil linhas por vinte e seis colunas e nenhum dado.
 * Imprimir a grade inteira gastaria dezenas de páginas em branco — e o usuário
 * descobriria isso na bandeja da impressora.
 */
export function usedBounds(sheet: Sheet): PrintBounds {
  let rows = 0
  let columns = 0

  for (const reference of Object.keys(sheet.cells)) {
    const match = /^([A-Z]+)([0-9]+)$/.exec(reference)
    if (match === null) continue

    rows = Math.max(rows, Number(match[2]))
    columns = Math.max(columns, indexOfColumn(match[1]!) + 1)
  }

  return { rows, columns }
}

function indexOfColumn(letters: string): number {
  let index = 0
  for (const letter of letters) index = index * 26 + (letter.charCodeAt(0) - 64)
  return index - 1
}

/**
 * Corpo da planilha em HTML.
 *
 * As linhas congeladas viram `<thead>`. Não é enfeite: o navegador repete o
 * `<thead>` no topo de cada página impressa, então o cabeçalho da tabela
 * reaparece na página 4 sem que ninguém tenha de programar isso. É a mesma
 * intenção de quem congelou a linha na tela.
 */
export function buildSheetHtml(sheet: Sheet): string {
  const bounds = usedBounds(sheet)
  if (bounds.rows === 0 || bounds.columns === 0) {
    return '<p class="sheet-print__empty">Esta aba está vazia.</p>'
  }

  const widths = columnWidths(sheet, bounds)

  const frozen = Math.min(sheet.frozenRows, bounds.rows)
  const head = frozen > 0 ? `<thead>${rowsHtml(sheet, bounds, 0, frozen)}</thead>` : ''
  const body = `<tbody>${rowsHtml(sheet, bounds, frozen, bounds.rows)}</tbody>`

  return `<table class="sheet-print" style="font-size:${fontSize(bounds.columns)}pt"><colgroup>${widths}</colgroup>${head}${body}</table>`
}

/**
 * Fonte menor conforme a planilha alarga — o "ajustar à página" do Excel.
 *
 * Doze colunas numa A4 em retrato dão cerca de cinquenta pixels cada. Em corpo
 * 11 nada cabe, e a quebra de linha começa a partir palavras no meio: "200" sai
 * como "20" e "0", que não é feio, é enganoso. Reduzir o corpo é o que evita
 * chegar nesse ponto, e é reversível pelo usuário — quem quiser o texto grande
 * põe a página em paisagem.
 */
function fontSize(columns: number): number {
  if (columns <= 8) return 11
  if (columns <= 12) return 9
  if (columns <= 18) return 8
  return 7
}

/**
 * Larguras em **proporção**, não em pixels.
 *
 * Em pixels, uma planilha mais larga que a página sai com a última coluna
 * cortada na margem — foi o que aconteceu na primeira versão, e o usuário só
 * descobriria depois de imprimir. Convertendo para porcentagem do total, a
 * tabela sempre cabe na folha e as colunas guardam a proporção que tinham na
 * tela, que é o que o usuário ajustou arrastando.
 */
function columnWidths(sheet: Sheet, bounds: PrintBounds): string {
  const pixels = Array.from(
    { length: bounds.columns },
    (_, column) => sheet.columnWidths[column] ?? DEFAULT_COLUMN_WIDTH,
  )

  const total = pixels.reduce((sum, width) => sum + width, 0)
  return pixels.map((width) => `<col style="width:${((width / total) * 100).toFixed(3)}%">`).join('')
}

function rowsHtml(sheet: Sheet, bounds: PrintBounds, from: number, to: number): string {
  const rows: string[] = []

  for (let row = from; row < to; row++) {
    const cells: string[] = []
    for (let column = 0; column < bounds.columns; column++) {
      cells.push(cellHtml(sheet.cells[cellRef(row, column)]))
    }

    const height = sheet.rowHeights[row]
    const style = height === undefined ? '' : ` style="height:${Math.round(height)}px"`
    rows.push(`<tr${style}>${cells.join('')}</tr>`)
  }

  return rows.join('')
}

function cellHtml(cell: Cell | undefined): string {
  const text = escapeHtml(formatCell(cell))
  const style = inlineStyle(cell)
  return style === '' ? `<td>${text}</td>` : `<td style="${style}">${text}</td>`
}

/**
 * Estilo direto na célula, e não classes.
 *
 * Cor e fundo são valores livres vindos do arquivo — uma classe por combinação
 * geraria uma folha de estilo do tamanho da planilha. Como este HTML existe por
 * alguns segundos dentro de uma janela oculta, o custo de estilo repetido é
 * pago em memória e não em manutenção.
 */
function inlineStyle(cell: Cell | undefined): string {
  const style = cell?.style
  if (style === undefined) return ''

  const parts: string[] = []
  if (style.bold === true) parts.push('font-weight:700')
  if (style.italic === true) parts.push('font-style:italic')
  if (style.underline === true) parts.push('text-decoration:underline')
  if (style.color !== undefined) parts.push(`color:${cssColor(style.color)}`)
  if (style.background !== undefined) parts.push(`background:${cssColor(style.background)}`)

  // Só o alinhamento escolhido, sem regra própria para número. O Excel joga
  // número à direita por conta própria, e a tentação de fazer o mesmo aqui é
  // grande — mas a grade **não** faz isso, e imprimir diferente do que está na
  // tela quebra a única promessa que a impressão tem. Se um dia isso mudar,
  // muda nos dois lugares.
  if (style.align !== undefined) parts.push(`text-align:${style.align}`)

  for (const side of style.borders ?? []) {
    parts.push(`border-${side}:1px solid #333`)
  }

  return parts.join(';')
}

/**
 * Cor só em `#rrggbb`.
 *
 * O valor vem de um documento, que é dado não confiável: sem esta trava, uma
 * cor como `red;background:url(...)` sairia do atributo `style` e viraria outra
 * declaração. Aqui isso só renderizaria algo estranho, mas a disciplina é a
 * mesma em todo lugar onde conteúdo de arquivo vira marcação.
 */
function cssColor(value: string): string {
  return /^#[0-9a-fA-F]{6}$/.test(value) ? value : 'inherit'
}

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/**
 * Folha de estilo da planilha impressa.
 *
 * `break-inside: avoid` na linha impede que uma linha alta seja cortada ao meio
 * pela quebra de página, que é o defeito mais visível de tabela impressa.
 */
export const SHEET_PRINT_CSS = `
.sheet-print {
  border-collapse: collapse;
  table-layout: fixed;
  width: 100%;
  font-family: Calibri, Carlito, system-ui, sans-serif;
  /* O corpo real vem no atributo da tabela: depende da largura da planilha. */
  font-size: 11pt;
}
/* O texto quebra em vez de ser cortado: na tela dá para alargar a coluna, no
   papel não, e conteúdo escondido no papel é perda silenciosa.

   Mas quebra **entre palavras**, nunca dentro de uma: quebrar em qualquer ponto
   parte "200" em "20" e "0" numa coluna estreita — e um número partido em duas
   linhas não é feio, é enganoso. Uma palavra que não couber transborda a célula:
   visível e estranho, que é melhor que invisível e errado. */
.sheet-print td {
  border: 1px solid #d0d0d0;
  padding: 2px 5px;
  vertical-align: bottom;
  word-break: normal;
  overflow-wrap: break-word;
}
.sheet-print tr { break-inside: avoid; }
.sheet-print thead td { font-weight: 700; background: #f2f2f2; }
.sheet-print__empty { font-family: system-ui, sans-serif; color: #666; }
`
