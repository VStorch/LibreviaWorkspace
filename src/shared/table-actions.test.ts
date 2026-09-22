import { describe, expect, it } from 'vitest'
import { TABLE_ACTIONS, TableAction } from './table-actions.js'
import { MenuCommand } from './types.js'

describe('ações de tabela', () => {
  it('toda ação é também um comando de menu, com o mesmo nome', () => {
    // O `App` repassa ao editor pelo nome. Uma ação que não estivesse em
    // `MenuCommand` seria recusada pelo contrato zod do IPC, e o item do menu
    // "Tabela" simplesmente não faria nada.
    const commands = new Set<string>(Object.values(MenuCommand))
    for (const action of Object.values(TableAction)) expect(commands.has(action)).toBe(true)
  })

  it('a lista do menu cobre cada ação uma vez', () => {
    const ids = TABLE_ACTIONS.map((action) => action.id)
    expect(new Set(ids).size).toBe(ids.length)
    expect([...ids].sort()).toEqual(Object.values(TableAction).sort())
  })

  it('só inserir tabela vale fora de uma tabela', () => {
    // É o que tira do menu de contexto os itens que não teriam o que fazer.
    expect(TABLE_ACTIONS.filter((action) => !action.needsTable).map((action) => action.id)).toEqual([
      TableAction.Insert,
    ])
  })
})
