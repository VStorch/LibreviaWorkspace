/** O grid é um web component: a posição está nos atributos que ele deixa no DOM da célula, o contrato público dele. */

function attributeOf(element: Element | null, name: string): number | null {
  const owner = element?.closest(`[${name}]`)
  const value = owner?.getAttribute(name)
  if (value === null || value === undefined) return null

  const index = Number.parseInt(value, 10)
  return Number.isInteger(index) ? index : null
}

/**
 * `null` no cabeçalho de coluna e na área vazia. `composedPath()[0]`, e não
 * `event.target`, que com shadow DOM chegaria reescrito como o hospedeiro.
 */
export function gridPositionOf(event: MouseEvent): { row: number; column: number } | null {
  const deepest = event.composedPath()[0]
  const target = deepest instanceof Element ? deepest : null

  const row = attributeOf(target, 'data-rgrow')
  const column = attributeOf(target, 'data-rgcol')
  return row === null || column === null ? null : { row, column }
}
