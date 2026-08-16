/**
 * Alça de preenchimento: arrastar copia a **fórmula deslocada**, não o número.
 *
 * É a diferença entre uma planilha e uma tabela. Quem arrasta `=B2*C2` para
 * baixo espera `=B3*C3` na linha seguinte — copiar o resultado calculado daria
 * a mesma coluna repetindo o valor da primeira linha, e o erro só apareceria no
 * fechamento do mês, quando os números não batessem.
 *
 * O deslocamento é o mesmo do Excel: referência relativa acompanha, referência
 * com `$` fica onde está. Quem faz isso é `translateFormula`, que já existia
 * desde a Fase 6 e só não tinha quem a chamasse.
 */

import { normalizeRange, type Range } from './edit.js'
import { translateFormula } from './formula/adjust.js'
import { getCell, setCell, type Cell, type Sheet } from './model.js'

/**
 * Preenche o destino a partir da origem.
 *
 * `source` é o que estava selecionado quando o arraste começou; `target` é a
 * área toda, origem incluída. As células da origem não são reescritas — elas já
 * estão certas, e regravá-las custaria recálculo à toa.
 *
 * A repetição é cíclica, como no Excel: uma origem de duas linhas arrastada por
 * seis preenche o padrão três vezes.
 */
export function fillRange(sheet: Sheet, source: Range, target: Range): Sheet {
  const from = normalizeRange(source)
  const to = normalizeRange(target)

  const rows = from.toRow - from.fromRow + 1
  const columns = from.toColumn - from.fromColumn + 1
  if (rows <= 0 || columns <= 0) return sheet

  let updated = sheet

  for (let row = to.fromRow; row <= to.toRow; row++) {
    for (let column = to.fromColumn; column <= to.toColumn; column++) {
      // Dentro da origem não se mexe.
      if (row >= from.fromRow && row <= from.toRow && column >= from.fromColumn && column <= from.toColumn) {
        continue
      }

      // Módulo com correção de sinal: arrastar **para cima** dá diferença
      // negativa, e `%` em JavaScript devolve negativo nesse caso — o que
      // apontaria para fora da origem.
      const sourceRow = from.fromRow + ((((row - from.fromRow) % rows) + rows) % rows)
      const sourceColumn = from.fromColumn + ((((column - from.fromColumn) % columns) + columns) % columns)

      updated = setCell(
        updated,
        row,
        column,
        copyOf(getCell(updated, sourceRow, sourceColumn), row - sourceRow, column - sourceColumn),
      )
    }
  }

  return updated
}

/**
 * A célula de origem, deslocada para o destino.
 *
 * O estilo vem junto — no Excel a alça leva a formatação, e deixá-la para trás
 * faria uma coluna de moeda preenchida virar meia coluna de números crus.
 *
 * O valor calculado **não** vem: ele é da posição de origem. Quem preenche o
 * valor da célula nova é o recálculo, logo em seguida, que é quem sabe a ordem
 * certa de calcular.
 */
function copyOf(source: Cell | undefined, rowDelta: number, columnDelta: number): Cell {
  if (source === undefined) return {}

  const cell: Cell =
    source.formula === undefined
      ? { ...(source.value === undefined ? {} : { value: source.value }) }
      : { formula: translateFormula(source.formula, rowDelta, columnDelta) }

  return source.style === undefined ? cell : { ...cell, style: source.style }
}
