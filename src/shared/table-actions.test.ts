import { describe, expect, it } from 'vitest'
import { TABLE_ACTIONS, TableAction } from './table-actions.js'
import { MenuCommand } from './types.js'

describe('table actions', () => {
  it('every action is also a menu command, with the same name', () => {
    // `App` forwards to the editor by name. An action missing from `MenuCommand` would be refused
    // by the IPC zod contract, and the "Table" menu item would do nothing.
    const commands = new Set<string>(Object.values(MenuCommand))
    for (const action of Object.values(TableAction)) expect(commands.has(action)).toBe(true)
  })

  it('the menu list covers each action once', () => {
    const ids = TABLE_ACTIONS.map((action) => action.id)
    expect(new Set(ids).size).toBe(ids.length)
    expect([...ids].sort()).toEqual(Object.values(TableAction).sort())
  })

  it('only inserting a table applies outside a table', () => {
    // This is what removes from the context menu the items with nothing to act on.
    expect(TABLE_ACTIONS.filter((action) => !action.needsTable).map((action) => action.id)).toEqual([
      TableAction.Insert,
    ])
  })
})
