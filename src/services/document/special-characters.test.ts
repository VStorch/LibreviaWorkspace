import { describe, expect, it } from 'vitest'
import { SPECIAL_CHARACTER_GROUPS, allSpecialCharacters } from './special-characters.js'

describe('special characters catalog', () => {
  it('no character appears twice', () => {
    // A duplicate is the defect that slips by when editing the list by hand, and on screen it
    // becomes two identical buttons in different groups.
    const all = allSpecialCharacters().map((item) => item.char)
    expect(new Set(all).size).toBe(all.length)
  })

  it('every character has a name, because it is the name the screen reader announces', () => {
    const unnamed = allSpecialCharacters().filter((item) => item.name.trim() === '')
    expect(unnamed).toEqual([])
  })

  it('every group has a label and content', () => {
    for (const group of SPECIAL_CHARACTER_GROUPS) {
      expect(group.label).not.toBe('')
      expect(group.characters.length).toBeGreaterThan(0)
    }
  })

  it('carries what is sought in a Portuguese document', () => {
    const all = allSpecialCharacters().map((item) => item.char)
    // Em dash, curly quotes and the non-breaking space that keeps "R$" next to the amount: the
    // three reasons to open this picker.
    expect(all).toContain('—')
    expect(all).toContain('“')
    expect(all).toContain('\u00a0')
  })
})
