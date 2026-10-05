/**
 * The editor's, not Chromium's, which would scale the whole UI. Through `transform`, which does not
 * change the `offsetTop` pagination measures.
 */

export const MIN_ZOOM = 50
export const MAX_ZOOM = 200

/** The same steps as LibreOffice and Word. */
export const ZOOM_STEPS: readonly number[] = [50, 67, 75, 90, 100, 110, 125, 150, 175, 200]

export function clampZoom(percent: number): number {
  if (!Number.isFinite(percent)) return 100
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, Math.round(percent)))
}

/** From a value off the ladder (fit width), the first one above it. */
export function zoomIn(percent: number): number {
  return ZOOM_STEPS.find((step) => step > percent) ?? MAX_ZOOM
}

export function zoomOut(percent: number): number {
  return [...ZOOM_STEPS].reverse().find((step) => step < percent) ?? MIN_ZOOM
}

/** With room on each side for the paper shadow. */
export function fitWidthZoom(availablePx: number, pageWidthPx: number, marginPx = 24): number {
  if (pageWidthPx <= 0 || availablePx <= 0) return 100
  return clampZoom(Math.floor(((availablePx - 2 * marginPx) / pageWidthPx) * 100))
}
