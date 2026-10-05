import { describe, expect, it } from 'vitest'
import { recalculate } from './formula/recalc.js'
import { createSheet, getCell, setCell, type Sheet, type WorkbookModel } from './model.js'
import { applyStructuralChange, isNameTaken, nextSheetName } from './structure.js'

/** Column A with 1, 2, 3 and their sum in C1. */
function planilha(name = 'Plan1'): Sheet {
  let sheet = createSheet(name)
  sheet = setCell(sheet, 0, 0, { value: 1 })
  sheet = setCell(sheet, 1, 0, { value: 2 })
  sheet = setCell(sheet, 2, 0, { value: 3 })
  sheet = setCell(sheet, 0, 2, { formula: '=SOMA(A1:A3)', value: 6 })
  return sheet
}

const pasta = (...sheets: Sheet[]): WorkbookModel => ({ sheets, activeSheet: 0 })

describe('inserir linha', () => {
  it('estica o intervalo da fórmula junto com os dados', () => {
    // Without the adjustment the data moved down and the formula kept summing the old place, giving
    // a wrong total without warning.
    const depois = applyStructuralChange(pasta(planilha()), 0, { kind: 'insertRows', at: 1, count: 1 })

    expect(getCell(depois.sheets[0]!, 0, 2)?.formula).toBe('=SOMA(A1:A4)')
  })

  it('e o total continua certo depois de recalcular', () => {
    let pastaAtual = applyStructuralChange(pasta(planilha()), 0, { kind: 'insertRows', at: 1, count: 1 })
    pastaAtual = recalculate(pastaAtual)

    expect(getCell(pastaAtual.sheets[0]!, 0, 2)?.value).toBe(6)
  })

  it('preenchendo a linha nova, o total acompanha', () => {
    let pastaAtual = applyStructuralChange(pasta(planilha()), 0, { kind: 'insertRows', at: 1, count: 1 })
    const comValor = setCell(pastaAtual.sheets[0]!, 1, 0, { value: 10 })
    pastaAtual = recalculate({ ...pastaAtual, sheets: [comValor] })

    expect(getCell(pastaAtual.sheets[0]!, 0, 2)?.value).toBe(16)
  })
})

describe('excluir linha', () => {
  // Row 2 is deleted: row 1 takes along the formula's own cell, which lives in it.
  it('encolhe o intervalo em vez de quebrar a fórmula', () => {
    const depois = applyStructuralChange(pasta(planilha()), 0, { kind: 'deleteRows', at: 1, count: 1 })

    expect(getCell(depois.sheets[0]!, 0, 2)?.formula).toBe('=SOMA(A1:A2)')
  })

  it('o total bate com o que sobrou', () => {
    const depois = recalculate(
      applyStructuralChange(pasta(planilha()), 0, { kind: 'deleteRows', at: 1, count: 1 }),
    )

    expect(getCell(depois.sheets[0]!, 0, 2)?.value).toBe(4)
  })

  it('excluir a célula apontada vira #REF!', () => {
    let sheet = createSheet('Plan1')
    sheet = setCell(sheet, 4, 0, { value: 9 })
    sheet = setCell(sheet, 0, 1, { formula: '=A5' })

    const depois = applyStructuralChange(pasta(sheet), 0, { kind: 'deleteRows', at: 4, count: 1 })

    expect(getCell(depois.sheets[0]!, 0, 1)?.formula).toBe('=#REF!')
  })
})

describe('colunas', () => {
  it('inserir empurra a referência para a direita', () => {
    const depois = applyStructuralChange(pasta(planilha()), 0, { kind: 'insertColumns', at: 0, count: 1 })

    expect(getCell(depois.sheets[0]!, 0, 3)?.formula).toBe('=SOMA(B1:B3)')
  })
})

describe('entre planilhas', () => {
  it('a fórmula da outra aba acompanha a linha inserida aqui', () => {
    // Why the operation belongs to the workbook and not the sheet: otherwise =Dados!A5 would keep
    // pointing where the data no longer is.
    let resumo = createSheet('Resumo')
    resumo = setCell(resumo, 0, 0, { formula: '=Dados!A5' })

    const depois = applyStructuralChange(pasta(planilha('Dados'), resumo), 0, {
      kind: 'insertRows',
      at: 0,
      count: 1,
    })

    expect(getCell(depois.sheets[1]!, 0, 0)?.formula).toBe('=Dados!A6')
  })

  it('a referência sem nome da outra aba não se mexe', () => {
    // =A5 on "Resumo" points to "Resumo", not to "Dados".
    let resumo = createSheet('Resumo')
    resumo = setCell(resumo, 0, 0, { formula: '=A5' })

    const depois = applyStructuralChange(pasta(planilha('Dados'), resumo), 0, {
      kind: 'insertRows',
      at: 0,
      count: 1,
    })

    expect(getCell(depois.sheets[1]!, 0, 0)?.formula).toBe('=A5')
  })
})

describe('identidade', () => {
  it('planilha sem fórmula volta como o mesmo objeto', () => {
    const semFormula = setCell(createSheet('Outra'), 0, 0, { value: 1 })
    const antes = pasta(planilha(), semFormula)
    const depois = applyStructuralChange(antes, 0, { kind: 'insertRows', at: 0, count: 1 })

    expect(depois.sheets[1]).toBe(semFormula)
  })

  it('operação sem efeito devolve a pasta intacta', () => {
    const antes = pasta(planilha())

    expect(applyStructuralChange(antes, 0, { kind: 'deleteRows', at: 5000, count: 1 })).toBe(antes)
  })
})

describe('nome da próxima aba', () => {
  it('segue a contagem quando os nomes são os padrões', () => {
    expect(nextSheetName(pasta(planilha('Planilha1')))).toBe('Planilha2')
  })

  it('pula o nome já usado em vez de repeti-lo', () => {
    // Someone who deleted Planilha2 and created another would have two with the same name, and
    // `=Planilha2!A1` would no longer have a single target.
    const atual = pasta(planilha('Planilha1'), planilha('Planilha3'))
    expect(nextSheetName(atual)).toBe('Planilha4')
  })

  it('reconhece o nome tomado por outra aba, e não pela própria', () => {
    const atual = pasta(planilha('Dados'), planilha('Resumo'))
    expect(isNameTaken(atual, 'Dados', 1)).toBe(true)
    expect(isNameTaken(atual, 'Dados', 0)).toBe(false)
  })
})
