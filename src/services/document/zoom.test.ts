import { describe, expect, it } from 'vitest'
import { clampZoom, fitWidthZoom, zoomIn, zoomOut } from './zoom.js'

describe('zoom da tela', () => {
  it('amplia e reduz pelos degraus, sem passar dos limites', () => {
    expect(zoomIn(100)).toBe(110)
    expect(zoomOut(100)).toBe(90)
    expect(zoomIn(200)).toBe(200)
    expect(zoomOut(50)).toBe(50)
  })

  it('de um valor ajustado à largura, vai ao degrau vizinho', () => {
    expect(zoomIn(113)).toBe(125)
    expect(zoomOut(113)).toBe(110)
  })

  it('ajustar à largura cabe a folha e respeita 50–200 %', () => {
    expect(fitWidthZoom(1648, 800)).toBe(200)
    expect(fitWidthZoom(848, 800)).toBe(100)
    expect(fitWidthZoom(300, 800)).toBe(50)
  })

  it('valor inválido volta a 100 %', () => {
    expect(clampZoom(Number.NaN)).toBe(100)
    expect(clampZoom(1000)).toBe(200)
  })
})
