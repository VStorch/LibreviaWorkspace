/**
 * O `Tab` circula entre os elementos de `selector` dentro do painel. Sem isso,
 * sairia para o texto ou para a barra, e o `Escape` não fecharia mais.
 */
export function cycleFocus(panel: HTMLElement | null, event: React.KeyboardEvent, selector: string): void {
  const stops = [...(panel?.querySelectorAll<HTMLElement>(selector) ?? [])]
  if (stops.length === 0) return

  const current = stops.indexOf(event.target as HTMLElement)
  const next = stops[(current + (event.shiftKey ? -1 : 1) + stops.length) % stops.length]
  if (next === undefined) return

  event.preventDefault()
  next.focus()
}
