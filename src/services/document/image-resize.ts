/**
 * In CSS pixels, the same as the node attributes; the writer converts to EMU (`ImageWriter.cs`).
 */

export const ResizeHandle = {
  NorthWest: 'nw',
  North: 'n',
  NorthEast: 'ne',
  East: 'e',
  SouthEast: 'se',
  South: 's',
  SouthWest: 'sw',
  West: 'w',
} as const
export type ResizeHandle = (typeof ResizeHandle)[keyof typeof ResizeHandle]

export const RESIZE_HANDLES: readonly ResizeHandle[] = Object.values(ResizeHandle)

/** A corner handle changes both measures; an edge handle, one. */
export function isCornerHandle(handle: ResizeHandle): boolean {
  return handle.length === 2
}

export interface ImageSize {
  readonly width: number
  readonly height: number
}

/** Below this the eight handles overlap and there is no way to grow back. */
export const MIN_IMAGE_PX = 24

export interface ResizeRequest {
  readonly handle: ResizeHandle
  /**
   * The size when the gesture started, not the previous frame's: the drag is measured from the
   * start.
   */
  readonly start: ImageSize
  readonly deltaX: number
  readonly deltaY: number
  /** Locked by default on corners and released by `Shift`, as in Word; never on edges. */
  readonly keepProportion: boolean
  /** The text column width. The image does not exceed it, as in Word. */
  readonly maxWidth: number
}

/**
 * Integers: an attribute like `342.7188` would change every frame from noise and pagination would
 * remeasure.
 */
export function resizedImage(request: ResizeRequest): ImageSize {
  const { handle, start, deltaX, deltaY, maxWidth } = request

  const horizontal = axisDirection(handle, 'e', 'w')
  const vertical = axisDirection(handle, 's', 'n')

  const ratio = start.height > 0 && start.width > 0 ? start.height / start.width : 0
  const locked = request.keepProportion && isCornerHandle(handle) && ratio > 0

  let width = start.width + horizontal * deltaX
  let height = start.height + vertical * deltaY

  // On a locked handle the width rules, the axis the column limits.
  if (locked) height = width * ratio
  if (horizontal === 0) width = locked ? height / ratio : start.width
  if (vertical === 0) height = locked ? width * ratio : start.height

  const bounded = withinBounds({ width, height }, Math.max(MIN_IMAGE_PX, maxWidth), locked ? ratio : null)
  return { width: Math.round(bounded.width), height: Math.round(bounded.height) }
}

function axisDirection(handle: ResizeHandle, forward: string, backward: string): number {
  if (handle.includes(forward)) return 1
  return handle.includes(backward) ? -1 : 0
}

/** Between the minimum and the column width; with the ratio locked, the other axis follows. */
function withinBounds(size: ImageSize, ceiling: number, ratio: number | null): ImageSize {
  let { width, height } = size
  if (width > ceiling || width < MIN_IMAGE_PX) {
    width = Math.min(ceiling, Math.max(MIN_IMAGE_PX, width))
    if (ratio !== null) height = width * ratio
  }
  if (height < MIN_IMAGE_PX) {
    height = MIN_IMAGE_PX
    if (ratio !== null) width = Math.min(ceiling, height / ratio)
  }
  return { width, height }
}

/**
 * When the column shrinks; otherwise the writer would shrink the image on its own and the screen
 * would lie.
 */
export function fittedImage(size: ImageSize, maxWidth: number): ImageSize {
  if (size.width <= maxWidth || size.width <= 0) return size

  const ceiling = Math.max(MIN_IMAGE_PX, maxWidth)
  return {
    width: Math.round(ceiling),
    height: Math.max(1, Math.round((size.height * ceiling) / size.width)),
  }
}
