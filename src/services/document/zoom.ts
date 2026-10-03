/**
 * Do editor, e não do Chromium, que aumentaria a interface inteira. Por
 * `transform`, que não muda o `offsetTop` que a paginação mede.
 */

export const MIN_ZOOM = 50
export const MAX_ZOOM = 200

/** Os mesmos degraus do LibreOffice e do Word. */
export const ZOOM_STEPS: readonly number[] = [50, 67, 75, 90, 100, 110, 125, 150, 175, 200]

export function clampZoom(percent: number): number {
  if (!Number.isFinite(percent)) return 100
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, Math.round(percent)))
}

/** O próximo degrau acima; de um valor fora da escada (o de ajustar), o primeiro acima dele. */
export function zoomIn(percent: number): number {
  return ZOOM_STEPS.find((step) => step > percent) ?? MAX_ZOOM
}

export function zoomOut(percent: number): number {
  return [...ZOOM_STEPS].reverse().find((step) => step < percent) ?? MIN_ZOOM
}

/** Com folga de cada lado para a sombra do papel. */
export function fitWidthZoom(availablePx: number, pageWidthPx: number, marginPx = 24): number {
  if (pageWidthPx <= 0 || availablePx <= 0) return 100
  return clampZoom(Math.floor(((availablePx - 2 * marginPx) / pageWidthPx) * 100))
}
