import { describe, expect, it } from 'vitest'
import { nextMenuIndex } from './menu-focus.js'

describe('menu arrow-key focus', () => {
  it('moves down and up in a loop', () => {
    expect(nextMenuIndex(0, 3, 'ArrowDown')).toBe(1)
    expect(nextMenuIndex(2, 3, 'ArrowDown')).toBe(0)
    expect(nextMenuIndex(0, 3, 'ArrowUp')).toBe(2)
  })

  it('with no focused item, arrow down goes to the first and arrow up to the last', () => {
    expect(nextMenuIndex(-1, 3, 'ArrowDown')).toBe(0)
    expect(nextMenuIndex(-1, 3, 'ArrowUp')).toBe(2)
  })

  it('Home and End go to the ends; other keys and an empty menu do not move', () => {
    expect(nextMenuIndex(1, 3, 'Home')).toBe(0)
    expect(nextMenuIndex(1, 3, 'End')).toBe(2)
    expect(nextMenuIndex(1, 3, 'a')).toBeNull()
    expect(nextMenuIndex(-1, 0, 'ArrowDown')).toBeNull()
  })
})
