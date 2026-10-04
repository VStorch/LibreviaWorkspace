import { describe, expect, it } from 'vitest'
import { mmToInches, mmToPx, pxToMm, pxToPt, twipsToMm, twipsToPx } from './units.js'

describe('mmToInches', () => {
  it.each([
    [25.4, 1],
    [12.7, 0.5],
    [0, 0],
  ])('converte %i mm em %f polegada', (mm, inches) => {
    expect(mmToInches(mm)).toBeCloseTo(inches, 10)
  })
})

describe('conversões de unidade', () => {
  it('uma polegada vale 25,4 mm, 96 px, 72 pt e 1440 twips', () => {
    expect(mmToPx(25.4)).toBeCloseTo(96, 10)
    expect(pxToMm(96)).toBeCloseTo(25.4, 10)
    expect(pxToPt(96)).toBe(72)
    expect(twipsToMm(1440)).toBeCloseTo(25.4, 10)
    expect(twipsToPx(1440)).toBe(96)
  })
})
