import { describe, expect, it } from 'vitest'
import { nextMenuIndex } from './menu-focus.js'

describe('foco nas setas do menu', () => {
  it('desce e sobe em círculo', () => {
    expect(nextMenuIndex(0, 3, 'ArrowDown')).toBe(1)
    expect(nextMenuIndex(2, 3, 'ArrowDown')).toBe(0)
    expect(nextMenuIndex(0, 3, 'ArrowUp')).toBe(2)
  })

  it('sem item focado, a seta para baixo vai ao primeiro e a para cima ao último', () => {
    expect(nextMenuIndex(-1, 3, 'ArrowDown')).toBe(0)
    expect(nextMenuIndex(-1, 3, 'ArrowUp')).toBe(2)
  })

  it('Home e End vão às pontas; outras teclas e menu vazio não movem', () => {
    expect(nextMenuIndex(1, 3, 'Home')).toBe(0)
    expect(nextMenuIndex(1, 3, 'End')).toBe(2)
    expect(nextMenuIndex(1, 3, 'a')).toBeNull()
    expect(nextMenuIndex(-1, 0, 'ArrowDown')).toBeNull()
  })
})
