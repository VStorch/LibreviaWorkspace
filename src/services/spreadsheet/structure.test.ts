import { describe, expect, it } from 'vitest'
import { recalculate } from './formula/recalc.js'
import { createSheet, getCell, setCell, type Sheet, type WorkbookModel } from './model.js'
import { applyStructuralChange, isNameTaken, nextSheetName } from './structure.js'

/** Coluna A com 1, 2, 3 e a soma delas em C1. */
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
    // O defeito que a Fase 5 deixou em aberto: os dados desciam e a fórmula
    // continuava somando o lugar antigo, dando um total errado sem aviso.
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
  // A exclusão é da linha 2: a linha 1 leva junto a própria célula da fórmula,
  // que mora nela.
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
    // O motivo de a operação ser da pasta e não da planilha: sem isso,
    // =Dados!A5 continuaria apontando para onde o dado não está mais.
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
    // =A5 em "Resumo" aponta para "Resumo", e não para "Dados".
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
    // Quem apagou a Planilha2 e criou outra teria duas com o mesmo nome, e
    // `=Planilha2!A1` deixaria de ter destino único.
    const atual = pasta(planilha('Planilha1'), planilha('Planilha3'))
    expect(nextSheetName(atual)).toBe('Planilha4')
  })

  it('reconhece o nome tomado por outra aba, e não pela própria', () => {
    const atual = pasta(planilha('Dados'), planilha('Resumo'))
    expect(isNameTaken(atual, 'Dados', 1)).toBe(true)
    expect(isNameTaken(atual, 'Dados', 0)).toBe(false)
  })
})
