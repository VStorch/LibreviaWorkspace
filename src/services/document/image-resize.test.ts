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

describe('resizing an image by its handles', () => {
  it('the bottom right corner grows in both dimensions', () => {
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

  it('the top left corner grows when the pointer moves backwards', () => {
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

  it('with the ratio unlocked the two axes move independently', () => {
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

  it('an edge handle changes one dimension only, even with the ratio requested', () => {
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

  it('the bottom handle changes only the height', () => {
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

  it('the image does not exceed the text column', () => {
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

  it('the image does not shrink until it cannot be grabbed back', () => {
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

  it('the size comes out in whole pixels', () => {
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

  it('eight handles: four corners and four edges', () => {
    expect(RESIZE_HANDLES).toHaveLength(8)
    expect(RESIZE_HANDLES.filter(isCornerHandle)).toHaveLength(4)
  })
})

describe('an image that stopped fitting', () => {
  it('a shrinking column shrinks the image proportionally', () => {
    expect(fittedImage({ width: 800, height: 600 }, 400)).toEqual({ width: 400, height: 300 })
  })

  it('an image that already fits stays as it is', () => {
    const size = { width: 300, height: 200 }
    expect(fittedImage(size, 400)).toBe(size)
  })
})
