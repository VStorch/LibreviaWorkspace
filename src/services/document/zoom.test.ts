import { describe, expect, it } from 'vitest'
import { clampZoom, fitWidthZoom, zoomIn, zoomOut } from './zoom.js'

describe('screen zoom', () => {
  it('zooms in and out by steps, within the limits', () => {
    expect(zoomIn(100)).toBe(110)
    expect(zoomOut(100)).toBe(90)
    expect(zoomIn(200)).toBe(200)
    expect(zoomOut(50)).toBe(50)
  })

  it('from a fit-to-width value, goes to the neighboring step', () => {
    expect(zoomIn(113)).toBe(125)
    expect(zoomOut(113)).toBe(110)
  })

  it('fit to width fits the sheet and respects 50–200 %', () => {
    expect(fitWidthZoom(1648, 800)).toBe(200)
    expect(fitWidthZoom(848, 800)).toBe(100)
    expect(fitWidthZoom(300, 800)).toBe(50)
  })

  it('an invalid value goes back to 100 %', () => {
    expect(clampZoom(Number.NaN)).toBe(100)
    expect(clampZoom(1000)).toBe(200)
  })
})
