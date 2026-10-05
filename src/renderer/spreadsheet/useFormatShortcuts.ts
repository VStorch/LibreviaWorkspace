import { useEffect, useRef } from 'react'

const SHORTCUTS: Record<string, 'bold' | 'italic' | 'underline' | undefined> = {
  b: 'bold',
  i: 'italic',
  u: 'underline',
}

/**
 * On `document`: focus lives inside the grid, a web component, and the event does not always bubble
 * to React.
 */
export function useFormatShortcuts(onToggle: (key: 'bold' | 'italic' | 'underline') => void): void {
  // Mounted once: the grid renders on every key press.
  const toggle = useRef(onToggle)
  toggle.current = onToggle

  useEffect(() => {
    const shortcut = (event: KeyboardEvent): void => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey) return

      // In a text field, the shortcut belongs to the field.
      if (event.target instanceof HTMLElement && event.target.closest('input, textarea') !== null) return

      const key = SHORTCUTS[event.key.toLowerCase()]
      if (key === undefined) return

      event.preventDefault()
      toggle.current(key)
    }

    document.addEventListener('keydown', shortcut)
    return () => document.removeEventListener('keydown', shortcut)
  }, [])
}
