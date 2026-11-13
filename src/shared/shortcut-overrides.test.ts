import { describe, expect, it } from 'vitest'
import { REMAPPABLE_SHORTCUTS } from './commands/remappable-shortcuts.js'
import {
  ShortcutProblem,
  applyShortcutOverrides,
  isRemappable,
  keyText,
  parseKeyText,
  shortcutKeyOf,
  shortcutProblem,
} from './shortcut-overrides.js'
import { SHORTCUTS, acceleratorOf, canonicalKeyOf } from './shortcuts.js'

const press = (code: string, modifiers: { ctrl?: true; shift?: true; alt?: true; meta?: true } = {}) => ({
  code,
  ctrlKey: modifiers.ctrl === true,
  metaKey: modifiers.meta === true,
  shiftKey: modifiers.shift === true,
  altKey: modifiers.alt === true,
})

describe('custom shortcuts', () => {
  it('moves a menu command to the chosen key and leaves the rest as they were', () => {
    const table = applyShortcutOverrides(SHORTCUTS, { save: 'Mod+Shift+K' })
    expect(acceleratorOf(table.save)).toBe('CmdOrCtrl+Shift+K')
    expect(acceleratorOf(table.open)).toBe('CmdOrCtrl+O')
  })

  it('lets two commands swap keys', () => {
    const table = applyShortcutOverrides(SHORTCUTS, { save: 'Mod+O', open: 'Mod+S' })
    expect(acceleratorOf(table.save)).toBe('CmdOrCtrl+O')
    expect(acceleratorOf(table.open)).toBe('CmdOrCtrl+S')
  })

  it('drops a hand-edited choice that takes an editor key, a system key or another command’s', () => {
    const table = applyShortcutOverrides(SHORTCUTS, { save: 'Mod+B', print: 'Mod+C', open: 'Mod+N' })
    expect(acceleratorOf(table.save)).toBe('CmdOrCtrl+S')
    expect(acceleratorOf(table.print)).toBe('CmdOrCtrl+P')
    expect(acceleratorOf(table.open)).toBe('CmdOrCtrl+O')
  })

  it('ignores editor keys, the keys the sheet hears and unknown ids', () => {
    const table = applyShortcutOverrides(SHORTCUTS, {
      bold: 'Mod+Shift+K',
      nextSheet: 'Mod+Shift+K',
      nothing: 'Mod+Shift+K',
    })
    expect(table).toEqual(SHORTCUTS)
    expect(isRemappable('alignLeft')).toBe(false)
    expect(isRemappable('constructor')).toBe(false)
  })

  it('offers every key the menu registers, each once, and no editor key', () => {
    const ids = REMAPPABLE_SHORTCUTS.map((entry) => entry.id)
    expect(new Set(ids).size).toBe(ids.length)
    expect(ids).toContain('save')
    expect(ids).toContain('formattingMarks')
    expect(ids).not.toContain('pasteWithoutFormat')
    expect(ids.every((id) => SHORTCUTS[id].owner === 'menu')).toBe(true)
  })

  it('says why a key cannot be used', () => {
    expect(shortcutProblem(SHORTCUTS, 'save', { key: 'K' })).toEqual({ kind: ShortcutProblem.NeedsModifier })
    expect(shortcutProblem(SHORTCUTS, 'save', { mod: true, key: 'V' })).toEqual({
      kind: ShortcutProblem.Reserved,
    })
    expect(shortcutProblem(SHORTCUTS, 'save', { mod: true, key: 'B' })).toEqual({
      kind: ShortcutProblem.Taken,
      by: 'bold',
    })
    expect(shortcutProblem(SHORTCUTS, 'save', { mod: true, key: 'S' })).toBeNull()
    expect(shortcutProblem(SHORTCUTS, 'save', { key: 'F7' })).toBeNull()
  })

  it('reads the physical key, so Shift+1 is 1 and not !', () => {
    expect(shortcutKeyOf(press('Digit1', { ctrl: true, shift: true }))).toEqual({
      mod: true,
      shift: true,
      key: '1',
    })
    expect(shortcutKeyOf(press('KeyK', { meta: true }))).toEqual({ mod: true, key: 'K' })
    expect(shortcutKeyOf(press('ArrowUp', { ctrl: true }))).toEqual({ mod: true, key: 'Up' })
    expect(shortcutKeyOf(press('NumpadAdd', { ctrl: true }))).toEqual({ mod: true, key: 'numadd' })
    expect(shortcutKeyOf(press('F5'))).toEqual({ key: 'F5' })
    expect(shortcutKeyOf(press('ControlLeft', { ctrl: true }))).toBeNull()
  })

  it('stores a key in a form that reads back the same, named keys included', () => {
    for (const key of [
      { mod: true, shift: true, key: 'PageDown' },
      { key: 'F9' },
      { mod: true, key: '-' },
    ] as const) {
      expect(canonicalKeyOf(parseKeyText(keyText(key))!)).toBe(canonicalKeyOf(key))
    }
    expect(parseKeyText('Hyper+K')).toBeNull()
    expect(parseKeyText('Mod+')).toBeNull()
  })
})
