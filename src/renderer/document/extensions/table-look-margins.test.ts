import { describe, expect, it } from 'vitest'
import { cellMarginsCss } from './table-look.js'

describe('margem de célula da tabela', () => {
  it('twips viram pixels de CSS, na ordem do padding', () => {
    expect(cellMarginsCss('0 108 60 108')).toBe('0px 7.2px 4px 7.2px')
  })

  it('o que não for quatro medidas some, e a célula fica com a margem do modelo', () => {
    expect(cellMarginsCss(null)).toBeNull()
    expect(cellMarginsCss('0 108')).toBeNull()
    expect(cellMarginsCss('0 -5 0 108')).toBeNull()
  })
})
