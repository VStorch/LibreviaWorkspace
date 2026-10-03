/** Arrastar copia a fórmula deslocada (`=B2*C2` → `=B3*C3`), e não o resultado. */

import { normalizeRange, type Range } from './edit.js'
import { translateFormula } from './formula/adjust.js'
import { getCell, setCell, type Cell, type Sheet } from './model.js'

/**
 * `target` inclui a origem, que não é reescrita. A repetição é cíclica, como no
 * Excel: duas linhas arrastadas por seis repetem o padrão três vezes.
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
      if (row >= from.fromRow && row <= from.toRow && column >= from.fromColumn && column <= from.toColumn) {
        continue
      }

      // Arrastar para cima dá diferença negativa, e `%` em JavaScript a devolve negativa.
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

/** O estilo vem junto, como no Excel; o valor não, porque é o recálculo que o preenche. */
function copyOf(source: Cell | undefined, rowDelta: number, columnDelta: number): Cell {
  if (source === undefined) return {}

  const cell: Cell =
    source.formula === undefined
      ? { ...(source.value === undefined ? {} : { value: source.value }) }
      : { formula: translateFormula(source.formula, rowDelta, columnDelta) }

  return source.style === undefined ? cell : { ...cell, style: source.style }
}
