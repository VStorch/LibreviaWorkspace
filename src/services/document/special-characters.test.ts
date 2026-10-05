import { describe, expect, it } from 'vitest'
import { SPECIAL_CHARACTER_GROUPS, allSpecialCharacters } from './special-characters.js'

describe('catálogo de caracteres especiais', () => {
  it('nenhum caractere aparece duas vezes', () => {
    // A duplicate is the defect that slips by when editing the list by hand, and on screen it
    // becomes two identical buttons in different groups.
    const todos = allSpecialCharacters().map((item) => item.char)
    expect(new Set(todos).size).toBe(todos.length)
  })

  it('todo caractere tem nome, porque é o nome que o leitor de tela anuncia', () => {
    const semNome = allSpecialCharacters().filter((item) => item.name.trim() === '')
    expect(semNome).toEqual([])
  })

  it('todo grupo tem rótulo e conteúdo', () => {
    for (const grupo of SPECIAL_CHARACTER_GROUPS) {
      expect(grupo.label).not.toBe('')
      expect(grupo.characters.length).toBeGreaterThan(0)
    }
  })

  it('traz o que se procura num documento em português', () => {
    const todos = allSpecialCharacters().map((item) => item.char)
    // Em dash, curly quotes and the non-breaking space that keeps "R$" next to the amount: the
    // three reasons to open this picker.
    expect(todos).toContain('—')
    expect(todos).toContain('“')
    expect(todos).toContain('\u00a0')
  })
})
