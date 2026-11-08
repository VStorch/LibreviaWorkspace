import { describe, expect, it } from 'vitest'
import { MESSAGES } from '../i18n/index.js'
import { SHORTCUTS, ShortcutOwner } from '../shortcuts.js'
import { MenuCommand } from '../types.js'
import { COMMANDS, type CommandMeta } from './registry.js'

const entries: readonly CommandMeta[] = Object.values(COMMANDS)

describe('command registry', () => {
  it('has one entry per menu command, under its own id', () => {
    expect(Object.keys(COMMANDS).sort()).toEqual(Object.values(MenuCommand).sort())
    for (const [id, meta] of Object.entries(COMMANDS)) expect(meta.id).toBe(id)
  })

  it('labels every command with a key the catalog has', () => {
    for (const meta of entries) expect(MESSAGES).toHaveProperty([meta.labelKey])
  })

  it('registers only menu-owned shortcuts as accelerators', () => {
    // The editor's keys reach the menu only shown: registering them would take them from the page.
    for (const meta of entries) {
      if (meta.shortcut === undefined) continue
      expect(SHORTCUTS).toHaveProperty([meta.shortcut])
      if (meta.registerAccelerator !== false) expect(SHORTCUTS[meta.shortcut].owner).toBe(ShortcutOwner.Menu)
    }
  })
})
