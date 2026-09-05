/**
 * Onde na grade o ponteiro caiu.
 *
 * O grid é um web component: a posição não vem no evento do React, e ler os
 * atributos que ele deixa no DOM da célula é o contrato público dele para
 * descobri-la.
 */

/** O índice que o grid deixa no DOM, ou `null` quando o elemento não o traz. */
function attributeOf(element: Element | null, name: string): number | null {
  const owner = element?.closest(`[${name}]`)
  const value = owner?.getAttribute(name)
  if (value === null || value === undefined) return null

  const index = Number.parseInt(value, 10)
  return Number.isInteger(index) ? index : null
}

/**
 * A célula sob o evento, ou `null` quando ele não caiu em nenhuma — cabeçalho
 * de coluna e área vazia não trazem posição.
 *
 * `composedPath()[0]` e não `event.target`: se um dia o grid passar a usar
 * shadow DOM o alvo chegaria aqui já reescrito como o elemento hospedeiro, sem
 * posição nenhuma.
 */
export function gridPositionOf(event: MouseEvent): { row: number; column: number } | null {
  const deepest = event.composedPath()[0]
  const target = deepest instanceof Element ? deepest : null

  const row = attributeOf(target, 'data-rgrow')
  const column = attributeOf(target, 'data-rgcol')
  return row === null || column === null ? null : { row, column }
}
