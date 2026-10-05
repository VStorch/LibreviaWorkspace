import { describe, expect, it } from 'vitest'
import { CellFormat, createSheet, getCell, setCell, type Sheet } from './model.js'
import {
  applyBorders,
  applyStyle,
  cellsIn,
  clearContents,
  deleteColumns,
  deleteRows,
  describeRange,
  insertColumns,
  insertRows,
  rangeContains,
  singleCell,
  toggleStyle,
  writeText,
} from './edit.js'

const range = { fromRow: 0, fromColumn: 0, toRow: 1, toColumn: 1 }

/** A1:B2 filled and a loose value in D4. */
function filled(): Sheet {
  let sheet = createSheet('S')
  sheet = setCell(sheet, 0, 0, { value: 'a' })
  sheet = setCell(sheet, 0, 1, { value: 'b' })
  sheet = setCell(sheet, 1, 0, { value: 'c' })
  sheet = setCell(sheet, 1, 1, { value: 'd' })
  sheet = setCell(sheet, 3, 3, { value: 'longe' })
  return sheet
}

describe('intervalo', () => {
  it('descreve célula única e intervalo', () => {
    expect(describeRange(singleCell(2, 1))).toBe('B3')
    expect(describeRange(range)).toBe('A1:B2')
  })

  it('normaliza intervalo selecionado de trás para frente', () => {
    // Dragging bottom-up is as common as the opposite.
    expect(describeRange({ fromRow: 5, fromColumn: 3, toRow: 1, toColumn: 1 })).toBe('B2:D6')
  })

  it('percorre todas as células', () => {
    expect([...cellsIn(range)]).toHaveLength(4)
  })
})

describe('formatação', () => {
  it('mescla com o estilo existente em vez de substituir', () => {
    // Making it bold must not erase the cell's existing background.
    const painted = applyStyle(filled(), singleCell(0, 0), { background: '#ffff00' })
    const bolded = applyStyle(painted, singleCell(0, 0), { bold: true })

    expect(getCell(bolded, 0, 0)?.style).toEqual({ background: '#ffff00', bold: true })
  })

  it('aplica a todas as células do intervalo', () => {
    const styled = applyStyle(filled(), range, { align: 'center' })

    for (const { row, column } of cellsIn(range)) {
      expect(getCell(styled, row, column)?.style?.align).toBe('center')
    }
    // Outside the range, nothing changes.
    expect(getCell(styled, 3, 3)?.style).toBeUndefined()
  })

  it('formata células vazias, para digitar já formatado', () => {
    const styled = applyStyle(createSheet('S'), singleCell(5, 5), { format: CellFormat.Currency })

    expect(getCell(styled, 5, 5)?.style?.format).toBe(CellFormat.Currency)
  })

  it('preserva o valor ao mudar o estilo', () => {
    const styled = applyStyle(filled(), singleCell(0, 0), { bold: true })

    expect(getCell(styled, 0, 0)?.value).toBe('a')
  })
})

describe('alternar atributo', () => {
  it('liga quando a seleção está mista', () => {
    // Half bold: the expectation is to turn everything on, not to flip each cell.
    const half = applyStyle(filled(), singleCell(0, 0), { bold: true })
    const toggled = toggleStyle(half, range, 'bold')

    for (const { row, column } of cellsIn(range)) {
      expect(getCell(toggled, row, column)?.style?.bold).toBe(true)
    }
  })

  it('desliga quando tudo já está ligado', () => {
    const all = applyStyle(filled(), range, { bold: true })
    const toggled = toggleStyle(all, range, 'bold')

    for (const { row, column } of cellsIn(range)) {
      expect(getCell(toggled, row, column)?.style?.bold).toBeUndefined()
    }
  })

  it('não deixa estilo vazio para trás', () => {
    // Toggling on and off must give the cell back as it was, or the file grows with empty styles on
    // every click.
    const sheet = createSheet('S')
    const on = toggleStyle(sheet, singleCell(0, 0), 'bold')
    const off = toggleStyle(on, singleCell(0, 0), 'bold')

    expect(Object.keys(off.cells)).toEqual([])
  })
})

describe('bordas', () => {
  it('aplica e remove', () => {
    const bordered = applyBorders(filled(), singleCell(0, 0), ['top', 'bottom'])
    expect(getCell(bordered, 0, 0)?.style?.borders).toEqual(['top', 'bottom'])

    const cleared = applyBorders(bordered, singleCell(0, 0), [])
    expect(getCell(cleared, 0, 0)?.style?.borders).toBeUndefined()
  })
})

describe('apagar conteúdo', () => {
  it('apaga o valor e preserva a formatação', () => {
    // That is what Delete does in a spreadsheet: clears the data, keeps the formatting.
    const styled = applyStyle(filled(), singleCell(0, 0), { background: '#eee' })
    const cleared = clearContents(styled, singleCell(0, 0))

    expect(getCell(cleared, 0, 0)?.value).toBeUndefined()
    expect(getCell(cleared, 0, 0)?.style?.background).toBe('#eee')
  })

  it('remove a célula que ficou sem nada', () => {
    const cleared = clearContents(filled(), singleCell(0, 0))
    expect(cleared.cells['A1']).toBeUndefined()
  })
})

describe('linhas e colunas', () => {
  it('insere linha deslocando o que vem depois', () => {
    const shifted = insertRows(filled(), 1)

    expect(getCell(shifted, 0, 0)?.value).toBe('a')
    expect(getCell(shifted, 2, 0)?.value).toBe('c')
    expect(getCell(shifted, 4, 3)?.value).toBe('longe')
  })

  it('exclui linha e puxa o resto para cima', () => {
    const shifted = deleteRows(filled(), 0)

    expect(getCell(shifted, 0, 0)?.value).toBe('c')
    expect(getCell(shifted, 2, 3)?.value).toBe('longe')
  })

  it('insere e exclui coluna', () => {
    const inserted = insertColumns(filled(), 0)
    expect(getCell(inserted, 0, 1)?.value).toBe('a')

    const deleted = deleteColumns(inserted, 0)
    expect(getCell(deleted, 0, 0)?.value).toBe('a')
  })

  it('desloca as larguras junto com as colunas', () => {
    // Otherwise inserting a column would leave the width on the wrong column.
    const sheet = { ...filled(), columnWidths: { 0: 200, 3: 60 } }
    const shifted = insertColumns(sheet, 0)

    expect(shifted.columnWidths[1]).toBe(200)
    expect(shifted.columnWidths[4]).toBe(60)
  })

  it('exclui várias linhas de uma vez', () => {
    const shifted = deleteRows(filled(), 0, 2)

    expect(shifted.cells['A1']).toBeUndefined()
    expect(getCell(shifted, 1, 3)?.value).toBe('longe')
  })

  it('cresce e encolhe a planilha junto', () => {
    // Otherwise inserting would push the last row out of the visible area: the data would vanish
    // from the screen and stay in the file.
    const sheet = filled()

    expect(insertRows(sheet, 0, 3).rowCount).toBe(sheet.rowCount + 3)
    expect(deleteRows(sheet, 0, 3).rowCount).toBe(sheet.rowCount - 3)
    expect(insertColumns(sheet, 0).columnCount).toBe(sheet.columnCount + 1)
    expect(deleteColumns(sheet, 0).columnCount).toBe(sheet.columnCount - 1)
  })

  it('não exclui além do fim da planilha', () => {
    const sheet = { ...createSheet('S'), rowCount: 10, columnCount: 4 }

    expect(deleteRows(sheet, 8, 500).rowCount).toBe(8)
    expect(deleteColumns(sheet, 3, 500).columnCount).toBe(3)
    // Past the end there is nothing to delete: the sheet comes back intact.
    expect(deleteRows(sheet, 10, 1)).toBe(sheet)
  })

  it('leva o congelamento junto quando a operação cai dentro dele', () => {
    const sheet = { ...createSheet('S'), frozenRows: 2, frozenColumns: 2 }

    expect(insertRows(sheet, 0).frozenRows).toBe(3)
    expect(deleteRows(sheet, 0).frozenRows).toBe(1)
    expect(insertColumns(sheet, 0).frozenColumns).toBe(3)
    // Below the frozen band, the freeze does not move.
    expect(insertRows(sheet, 2).frozenRows).toBe(2)
  })
})

describe('rangeContains', () => {
  it('aceita as bordas do retângulo e recusa o que está fora', () => {
    expect(rangeContains(range, 0, 0)).toBe(true)
    expect(rangeContains(range, 1, 1)).toBe(true)
    expect(rangeContains(range, 2, 1)).toBe(false)
    expect(rangeContains(range, 1, 2)).toBe(false)
  })

  it('vale para o intervalo arrastado de baixo para cima', () => {
    expect(rangeContains({ fromRow: 3, fromColumn: 3, toRow: 1, toColumn: 1 }, 2, 2)).toBe(true)
  })
})

describe('writeText', () => {
  it('guarda a fórmula e deixa o valor para o recálculo', () => {
    const sheet = writeText(createSheet('S'), 0, 0, '=SOMA(A2:A3)')
    expect(getCell(sheet, 0, 0)).toEqual({ formula: '=SOMA(A2:A3)' })
  })

  it('reconhece o formato do que foi digitado', () => {
    const sheet = writeText(createSheet('S'), 0, 0, 'R$ 10,50')
    expect(getCell(sheet, 0, 0)?.value).toBe(10.5)
    expect(getCell(sheet, 0, 0)?.style?.format).toBe(CellFormat.Currency)
  })

  it('não apaga o estilo que o usuário escolheu à mão', () => {
    const painted = setCell(createSheet('S'), 0, 0, { value: 1, style: { bold: true } })
    expect(writeText(painted, 0, 0, '2').cells['A1']).toEqual({ value: 2, style: { bold: true } })
  })
})
