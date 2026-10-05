import { describe, expect, it } from 'vitest'
import { createSheet, getCell, setCell, type Sheet, type WorkbookModel } from '../model.js'
import { FormulaError } from './errors.js'
import { recalculate } from './recalc.js'

/** A sheet from a map of reference → value or formula. */
function sheetWith(name: string, entries: Record<string, string | number>): Sheet {
  let sheet = createSheet(name)

  for (const [ref, entry] of Object.entries(entries)) {
    const match = /^([A-Z]+)(\d+)$/.exec(ref)!
    let column = 0
    for (const letter of match[1]!) column = column * 26 + (letter.charCodeAt(0) - 64)
    const row = Number(match[2]) - 1

    sheet =
      typeof entry === 'string' && entry.startsWith('=')
        ? setCell(sheet, row, column - 1, { formula: entry })
        : setCell(sheet, row, column - 1, { value: entry })
  }

  return sheet
}

function workbookWith(...sheets: Sheet[]): WorkbookModel {
  return { sheets, activeSheet: 0 }
}

function valueOf(workbook: WorkbookModel, ref: string, sheet = 0) {
  const match = /^([A-Z]+)(\d+)$/.exec(ref)!
  let column = 0
  for (const letter of match[1]!) column = column * 26 + (letter.charCodeAt(0) - 64)
  return getCell(workbook.sheets[sheet]!, Number(match[2]) - 1, column - 1)?.value
}

describe('ordem de cálculo', () => {
  it('calcula a cadeia de trás para frente, não na ordem do mapa', () => {
    // A1 depends on B1, which depends on C1. Computing in appearance order would give A1 the old
    // value of B1, and the error would only show the second time someone touched the sheet.
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=B1+1', B1: '=C1*2', C1: 5 })))

    expect(valueOf(done, 'C1')).toBe(5)
    expect(valueOf(done, 'B1')).toBe(10)
    expect(valueOf(done, 'A1')).toBe(11)
  })

  it('resolve cadeia longa sem estourar a pilha', () => {
    // A recursive walk would die here, and the user would see the app vanish without any
    // explanation. The map is built directly because `setCell` copies the whole sheet for each
    // cell, and five thousand copies would dominate the test.
    const cells: Record<string, { formula?: string; value?: number }> = { A1: { value: 1 } }
    for (let row = 2; row <= 5000; row++) cells[`A${row}`] = { formula: `=A${row - 1}+1` }

    const done = recalculate(workbookWith({ ...createSheet('P'), cells }))

    expect(valueOf(done, 'A5000')).toBe(5000)
  })

  it('propaga por intervalo', () => {
    const done = recalculate(
      workbookWith(sheetWith('P', { A1: 1, A2: '=A1*10', A3: '=A2*10', B1: '=SOMA(A1:A3)' })),
    )

    expect(valueOf(done, 'B1')).toBe(111)
  })
})

describe('referência circular', () => {
  it('marca as duas pontas de um ciclo direto', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=B1', B1: '=A1' })))

    expect(valueOf(done, 'A1')).toBe(FormulaError.Circular)
    expect(valueOf(done, 'B1')).toBe(FormulaError.Circular)
  })

  it('pega a célula que se referencia sozinha', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=A1+1' })))

    expect(valueOf(done, 'A1')).toBe(FormulaError.Circular)
  })

  it('pega o ciclo pela soma de um intervalo que contém a própria célula', () => {
    // The most common way to create a cycle unknowingly: dragging the sum into the column it sums.
    const done = recalculate(workbookWith(sheetWith('P', { A1: 1, A2: 2, A3: '=SOMA(A1:A3)' })))

    expect(valueOf(done, 'A3')).toBe(FormulaError.Circular)
  })

  it('quem depende de uma circular herda o erro, e não um valor velho', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=B1', B1: '=A1', C1: '=A1+1' })))

    expect(valueOf(done, 'C1')).toBe(FormulaError.Circular)
  })

  it('o resto da planilha continua calculando', () => {
    // A cycle in one corner must not bring down the rest of the calculation.
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=B1', B1: '=A1', D1: 2, D2: '=D1*3' })))

    expect(valueOf(done, 'D2')).toBe(6)
  })
})

describe('entre planilhas', () => {
  it('lê célula de outra aba', () => {
    const done = recalculate(
      workbookWith(sheetWith('Resumo', { A1: '=Dados!B2*2' }), sheetWith('Dados', { B2: 21 })),
    )

    expect(valueOf(done, 'A1')).toBe(42)
  })

  it('referência sem nome aponta para a planilha da própria fórmula', () => {
    // The trap: if the default were the active tab, the second tab's formula would read the first
    // one's cell and give a plausible, wrong number.
    const done = recalculate(
      workbookWith(sheetWith('Um', { A1: 100, B1: '=A1' }), sheetWith('Dois', { A1: 7, B1: '=A1' })),
    )

    expect(valueOf(done, 'B1', 0)).toBe(100)
    expect(valueOf(done, 'B1', 1)).toBe(7)
  })

  it('não diferencia maiúsculas no nome da aba', () => {
    const done = recalculate(
      workbookWith(sheetWith('Resumo', { A1: '=dados!B2' }), sheetWith('Dados', { B2: 9 })),
    )

    expect(valueOf(done, 'A1')).toBe(9)
  })

  it('aba que não existe vira #REF!', () => {
    const done = recalculate(workbookWith(sheetWith('Resumo', { A1: '=Sumida!B2' })))

    expect(valueOf(done, 'A1')).toBe(FormulaError.Ref)
  })

  it('detecta ciclo que atravessa abas', () => {
    const done = recalculate(
      workbookWith(sheetWith('Um', { A1: '=Dois!A1' }), sheetWith('Dois', { A1: '=Um!A1' })),
    )

    expect(valueOf(done, 'A1', 0)).toBe(FormulaError.Circular)
    expect(valueOf(done, 'A1', 1)).toBe(FormulaError.Circular)
  })
})

describe('o que o recálculo preserva', () => {
  it('mantém a fórmula ao lado do valor', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: 2, B1: '=A1*3' })))

    expect(getCell(done.sheets[0]!, 0, 1)).toEqual({ formula: '=A1*3', value: 6 })
  })

  it('devolve a mesma pasta quando não há fórmula nenhuma', () => {
    // Identity matters: React redraws the whole grid if the object changes.
    const workbook = workbookWith(sheetWith('P', { A1: 1, B1: 2 }))

    expect(recalculate(workbook)).toBe(workbook)
  })

  it('devolve a mesma pasta quando nenhum valor mudou', () => {
    const workbook = recalculate(workbookWith(sheetWith('P', { A1: 2, B1: '=A1*3' })))

    expect(recalculate(workbook)).toBe(workbook)
  })

  it('fórmula sobre célula vazia vale zero, e não some do arquivo', () => {
    // If an empty result erased the cell, the formula would go with it: the sparse map removes
    // cells without value, formula or style.
    const done = recalculate(workbookWith(sheetWith('P', { B1: '=A1' })))

    expect(getCell(done.sheets[0]!, 0, 1)).toEqual({ formula: '=A1', value: 0 })
  })

  it('fórmula que não fecha não derruba o recálculo', () => {
    // Only reachable with a hand-edited file; the UI refuses it first.
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=SOMA(', B1: 1, C1: '=B1+1' })))

    expect(valueOf(done, 'A1')).toBe(FormulaError.Value)
    expect(valueOf(done, 'C1')).toBe(2)
  })
})

describe('data injetada', () => {
  it('HOJE usa o relógio que recebe', () => {
    const done = recalculate(workbookWith(sheetWith('P', { A1: '=ANO(HOJE())' })), () => new Date(2030, 0, 1))

    expect(valueOf(done, 'A1')).toBe(2030)
  })
})
