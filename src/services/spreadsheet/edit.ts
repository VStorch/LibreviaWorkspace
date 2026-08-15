/**
 * Operações sobre um intervalo de células.
 *
 * Camada pura: recebe a planilha e devolve outra, sem tocar na original. É o
 * que permite testar formatação sem grid nenhum — e é onde o desfazer da Fase 8
 * vai se apoiar.
 */

import { cellRef, getCell, setCell, type BorderSide, type Cell, type CellStyle, type Sheet } from './model.js'

/** Um retângulo de células, em coordenadas base zero e inclusivas. */
export interface Range {
  readonly fromRow: number
  readonly fromColumn: number
  readonly toRow: number
  readonly toColumn: number
}

export function normalizeRange(range: Range): Range {
  return {
    fromRow: Math.min(range.fromRow, range.toRow),
    toRow: Math.max(range.fromRow, range.toRow),
    fromColumn: Math.min(range.fromColumn, range.toColumn),
    toColumn: Math.max(range.fromColumn, range.toColumn),
  }
}

export function singleCell(row: number, column: number): Range {
  return { fromRow: row, toRow: row, fromColumn: column, toColumn: column }
}

/** Referência exibível: "B3" para uma célula, "B3:D9" para um intervalo. */
export function describeRange(range: Range): string {
  const { fromRow, fromColumn, toRow, toColumn } = normalizeRange(range)
  const start = cellRef(fromRow, fromColumn)
  return fromRow === toRow && fromColumn === toColumn ? start : `${start}:${cellRef(toRow, toColumn)}`
}

export function* cellsIn(range: Range): Generator<{ row: number; column: number }> {
  const { fromRow, fromColumn, toRow, toColumn } = normalizeRange(range)
  for (let row = fromRow; row <= toRow; row++) {
    for (let column = fromColumn; column <= toColumn; column++) {
      yield { row, column }
    }
  }
}

/**
 * Aplica uma mudança de estilo a todas as células do intervalo.
 *
 * A alteração é **mesclada** com o estilo existente, não substituída: pôr uma
 * célula em negrito não pode apagar a cor de fundo que ela já tinha.
 */
export function applyStyle(sheet: Sheet, range: Range, change: Partial<CellStyle>): Sheet {
  let updated = sheet

  for (const { row, column } of cellsIn(range)) {
    const cell = getCell(updated, row, column) ?? {}
    const style = clean({ ...cell.style, ...change })

    updated = setCell(updated, row, column, toCell(cell, style))
  }

  return updated
}

/**
 * Liga ou desliga um atributo booleano no intervalo inteiro.
 *
 * A regra é a das planilhas e dos editores: se **tudo** já está ligado,
 * desliga; senão, liga tudo. Alternar célula a célula deixaria a seleção
 * misturada e o botão sem significado.
 */
export function toggleStyle(sheet: Sheet, range: Range, key: 'bold' | 'italic' | 'underline'): Sheet {
  const allOn = [...cellsIn(range)].every(
    ({ row, column }) => getCell(sheet, row, column)?.style?.[key] === true,
  )

  return applyStyle(sheet, range, { [key]: allOn ? undefined : true })
}

/**
 * Bordas do intervalo.
 *
 * `sides` vazio remove as bordas. Não há aqui a noção de "borda externa" — cada
 * célula recebe os mesmos lados, que é o que o modelo representa hoje.
 */
export function applyBorders(sheet: Sheet, range: Range, sides: readonly BorderSide[]): Sheet {
  return applyStyle(sheet, range, { borders: sides.length === 0 ? undefined : [...sides] })
}

/** Apaga o conteúdo, preservando a formatação — como a tecla Delete faz. */
export function clearContents(sheet: Sheet, range: Range): Sheet {
  let updated = sheet

  for (const { row, column } of cellsIn(range)) {
    const cell = getCell(updated, row, column)
    if (cell === undefined) continue

    updated = setCell(updated, row, column, cell.style === undefined ? {} : { style: cell.style })
  }

  return updated
}

/**
 * Insere linhas antes da posição indicada, deslocando o que vem depois.
 *
 * Percorre de baixo para cima ao mover: começar de cima sobrescreveria as
 * células ainda não movidas.
 */
export function insertRows(sheet: Sheet, at: number, count = 1): Sheet {
  return shiftRows(sheet, at, count)
}

export function deleteRows(sheet: Sheet, at: number, count = 1): Sheet {
  return shiftRows(sheet, at, -count)
}

export function insertColumns(sheet: Sheet, at: number, count = 1): Sheet {
  return shiftColumns(sheet, at, count)
}

export function deleteColumns(sheet: Sheet, at: number, count = 1): Sheet {
  return shiftColumns(sheet, at, -count)
}

function shiftRows(sheet: Sheet, at: number, delta: number): Sheet {
  const cells: Record<string, Cell> = {}

  for (const [ref, cell] of Object.entries(sheet.cells)) {
    const position = positionOf(ref)
    if (position === null) continue

    if (position.row < at) {
      cells[ref] = cell
      continue
    }

    // Linha excluída: a célula desaparece junto.
    if (delta < 0 && position.row < at - delta) continue

    cells[cellRef(position.row + delta, position.column)] = cell
  }

  return { ...sheet, cells, rowHeights: shiftDimensions(sheet.rowHeights, at, delta) }
}

function shiftColumns(sheet: Sheet, at: number, delta: number): Sheet {
  const cells: Record<string, Cell> = {}

  for (const [ref, cell] of Object.entries(sheet.cells)) {
    const position = positionOf(ref)
    if (position === null) continue

    if (position.column < at) {
      cells[ref] = cell
      continue
    }

    if (delta < 0 && position.column < at - delta) continue

    cells[cellRef(position.row, position.column + delta)] = cell
  }

  return { ...sheet, cells, columnWidths: shiftDimensions(sheet.columnWidths, at, delta) }
}

function shiftDimensions(sizes: Record<number, number>, at: number, delta: number): Record<number, number> {
  const shifted: Record<number, number> = {}

  for (const [key, size] of Object.entries(sizes)) {
    const index = Number(key)
    if (index < at) {
      shifted[index] = size
      continue
    }
    if (delta < 0 && index < at - delta) continue
    shifted[index + delta] = size
  }

  return shifted
}

function positionOf(ref: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)([0-9]+)$/.exec(ref)
  if (match === null) return null

  let column = 0
  for (const letter of match[1]!) column = column * 26 + (letter.charCodeAt(0) - 64)

  return { row: Number.parseInt(match[2]!, 10) - 1, column: column - 1 }
}

/** Remove as chaves indefinidas para que o estilo vazio seja de fato vazio. */
function clean(style: Record<string, unknown>): CellStyle {
  const result: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(style)) {
    if (value !== undefined) result[key] = value
  }
  return result as CellStyle
}

function toCell(cell: Cell, style: CellStyle): Cell {
  const base: Cell = {}
  const withValue = cell.value === undefined ? base : { ...base, value: cell.value }
  const withFormula = cell.formula === undefined ? withValue : { ...withValue, formula: cell.formula }

  return Object.keys(style).length === 0 ? withFormula : { ...withFormula, style }
}
