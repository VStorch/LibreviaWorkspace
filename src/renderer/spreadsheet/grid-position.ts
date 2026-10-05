/**
 * The grid is a web component: the position is in the attributes it leaves on the cell DOM, its
 * public contract.
 */

function attributeOf(element: Element | null, name: string): number | null {
  const owner = element?.closest(`[${name}]`)
  const value = owner?.getAttribute(name)
  if (value === null || value === undefined) return null

  const index = Number.parseInt(value, 10)
  return Number.isInteger(index) ? index : null
}

/**
 * `null` on the column header and the empty area. `composedPath()[0]`, not `event.target`, which
 * with shadow DOM would arrive retargeted to the host.
 */
export function gridPositionOf(event: MouseEvent): { row: number; column: number } | null {
  const deepest = event.composedPath()[0]
  const target = deepest instanceof Element ? deepest : null

  const row = attributeOf(target, 'data-rgrow')
  const column = attributeOf(target, 'data-rgcol')
  return row === null || column === null ? null : { row, column }
}
