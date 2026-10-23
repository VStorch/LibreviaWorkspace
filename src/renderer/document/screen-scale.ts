/**
 * Zoom is a `transform`: only what comes from the screen arrives multiplied. Whoever converts a
 * gesture into a measure divides by this. Read from the element, to hold with any transform along
 * the way.
 */
export function screenScaleOf(element: HTMLElement | null): number {
  if (element === null) return 1
  const width = layoutWidthOf(element)
  const screen = element.getBoundingClientRect().width
  return width > 0 && screen > 0 ? screen / width : 1
}

/**
 * Not `offsetWidth`, which rounds to whole pixels: the A4 sheet is 793.7 px wide, and 794 read as a
 * 0.04% zoom that pagination carried from one measurement into the next.
 */
function layoutWidthOf(element: HTMLElement): number {
  const style = getComputedStyle(element)
  const width = parseFloat(style.width)
  if (Number.isNaN(width)) return element.offsetWidth
  if (style.boxSizing === 'border-box') return width
  return [style.paddingLeft, style.paddingRight, style.borderLeftWidth, style.borderRightWidth].reduce(
    (total, side) => total + (parseFloat(side) || 0),
    width,
  )
}
