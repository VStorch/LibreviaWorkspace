/**
 * O zoom da tela, que é do editor e não do Chromium.
 *
 * O zoom do navegador aumentava a interface inteira — barras, menus, diálogos —
 * para aumentar a folha. Este aumenta só a pilha de papel, e por `transform`: a
 * paginação mede em pixels de CSS (`offsetTop`, `offsetHeight`), que a
 * transformação não toca. A folha medida em 100 % é a folha desenhada em 150 %,
 * e o corte não anda quando se amplia.
 */

export const MIN_ZOOM = 50
export const MAX_ZOOM = 200

/** Os degraus de ampliar e reduzir, os mesmos do LibreOffice e do Word. */
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

/**
 * O zoom em que a folha ocupa a largura disponível, com uma folga de cada lado
 * para a sombra do papel não encostar na borda.
 */
export function fitWidthZoom(availablePx: number, pageWidthPx: number, marginPx = 24): number {
  if (pageWidthPx <= 0 || availablePx <= 0) return 100
  return clampZoom(Math.floor(((availablePx - 2 * marginPx) / pageWidthPx) * 100))
}
