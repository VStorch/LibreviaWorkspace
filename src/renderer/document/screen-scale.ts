/**
 * Zoom is a `transform`: only what comes from the screen arrives multiplied. Whoever converts a
 * gesture into a measure divides by this. Read from the element, to hold with any transform along
 * the way.
 */
export function screenScaleOf(element: HTMLElement | null): number {
  if (element === null || element.offsetWidth <= 0) return 1
  const width = element.getBoundingClientRect().width
  return width > 0 ? width / element.offsetWidth : 1
}
