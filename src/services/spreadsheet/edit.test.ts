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
  singleCell,
  toggleStyle,
} from './edit.js'

const range = { fromRow: 0, fromColumn: 0, toRow: 1, toColumn: 1 }

/** Planilha com A1:B2 preenchidas e um valor solto em D4. */
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
    // Arrastar de baixo para cima é tão comum quanto o contrário.
    expect(describeRange({ fromRow: 5, fromColumn: 3, toRow: 1, toColumn: 1 })).toBe('B2:D6')
  })

  it('percorre todas as células', () => {
    expect([...cellsIn(range)]).toHaveLength(4)
  })
})

describe('formatação', () => {
  it('mescla com o estilo existente em vez de substituir', () => {
    // Pôr em negrito não pode apagar o fundo que a célula já tinha.
    const painted = applyStyle(filled(), singleCell(0, 0), { background: '#ffff00' })
    const bolded = applyStyle(painted, singleCell(0, 0), { bold: true })

    expect(getCell(bolded, 0, 0)?.style).toEqual({ background: '#ffff00', bold: true })
  })

  it('aplica a todas as células do intervalo', () => {
    const styled = applyStyle(filled(), range, { align: 'center' })

    for (const { row, column } of cellsIn(range)) {
      expect(getCell(styled, row, column)?.style?.align).toBe('center')
    }
    // Fora do intervalo, nada muda.
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
    // Metade em negrito: o esperado é ligar tudo, não inverter cada célula.
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
    // Ligar e desligar precisa devolver a célula ao que era, senão o arquivo
    // engorda com estilos vazios a cada clique.
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
    // É o que a tecla Delete faz numa planilha: limpa o dado, mantém o desenho.
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
    // Sem isso, inserir uma coluna deixaria a largura na coluna errada.
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
})
