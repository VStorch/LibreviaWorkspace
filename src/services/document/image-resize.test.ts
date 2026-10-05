import { describe, expect, it } from 'vitest'
import {
  MIN_IMAGE_PX,
  RESIZE_HANDLES,
  ResizeHandle,
  fittedImage,
  isCornerHandle,
  resizedImage,
} from './image-resize.js'

const start = { width: 400, height: 300 }
const wide = 624

describe('redimensionar imagem pelas alças', () => {
  it('o canto de baixo à direita cresce nas duas medidas', () => {
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start,
      deltaX: 100,
      deltaY: 0,
      keepProportion: true,
      maxWidth: wide,
    })

    // A locked ratio makes the height follow the width even when the pointer did not move
    // vertically: the horizontal axis rules.
    expect(size).toEqual({ width: 500, height: 375 })
  })

  it('o canto de cima à esquerda cresce quando o ponteiro vai para trás', () => {
    const size = resizedImage({
      handle: ResizeHandle.NorthWest,
      start,
      deltaX: -100,
      deltaY: -100,
      keepProportion: true,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: 500, height: 375 })
  })

  it('com a proporção solta os dois eixos andam sozinhos', () => {
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start,
      deltaX: 100,
      deltaY: -100,
      keepProportion: false,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: 500, height: 200 })
  })

  it('a alça de borda mexe numa medida só, mesmo com a proporção pedida', () => {
    // The middle handle exists to stretch one axis: locking the ratio on it would make it a corner.
    const size = resizedImage({
      handle: ResizeHandle.East,
      start,
      deltaX: 100,
      deltaY: 250,
      keepProportion: true,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: 500, height: 300 })
  })

  it('a alça de baixo mexe só na altura', () => {
    const size = resizedImage({
      handle: ResizeHandle.South,
      start,
      deltaX: 250,
      deltaY: 60,
      keepProportion: false,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: 400, height: 360 })
  })

  it('a imagem não passa da coluna de texto', () => {
    // Without the ceiling, the writer would shrink the image on its own when saving, and the screen
    // would no longer show what the file has.
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start,
      deltaX: 900,
      deltaY: 0,
      keepProportion: true,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: wide, height: 468 })
  })

  it('a imagem não encolhe até não dar para pegar de volta', () => {
    // Below the minimum the eight handles overlap, and growing back would only be possible by
    // undoing an unfinished drag.
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start,
      deltaX: -1000,
      deltaY: -1000,
      keepProportion: false,
      maxWidth: wide,
    })

    expect(size).toEqual({ width: MIN_IMAGE_PX, height: MIN_IMAGE_PX })
  })

  it('o tamanho sai em pixel inteiro', () => {
    // The number goes into the node attribute, and a value with decimals would change every frame
    // from floating-point noise, making pagination remeasure everything.
    const size = resizedImage({
      handle: ResizeHandle.SouthEast,
      start: { width: 401, height: 297 },
      deltaX: 3,
      deltaY: 0,
      keepProportion: true,
      maxWidth: wide,
    })

    expect(Number.isInteger(size.width)).toBe(true)
    expect(Number.isInteger(size.height)).toBe(true)
  })

  it('oito alças: quatro cantos e quatro bordas', () => {
    expect(RESIZE_HANDLES).toHaveLength(8)
    expect(RESIZE_HANDLES.filter(isCornerHandle)).toHaveLength(4)
  })
})

describe('imagem que deixou de caber', () => {
  it('a coluna que encolhe encolhe a imagem na proporção', () => {
    expect(fittedImage({ width: 800, height: 600 }, 400)).toEqual({ width: 400, height: 300 })
  })

  it('a imagem que já cabe fica como está', () => {
    const size = { width: 300, height: 200 }
    expect(fittedImage(size, 400)).toBe(size)
  })
})
