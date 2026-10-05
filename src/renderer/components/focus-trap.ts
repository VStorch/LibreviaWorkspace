/**
 * `Tab` cycles among `selector` elements inside the panel. Otherwise it would leave for the text or
 * the toolbar, and `Escape` would no longer close it.
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
