import { useEffect, useRef } from 'react'

const SHORTCUTS: Record<string, 'bold' | 'italic' | 'underline' | undefined> = {
  b: 'bold',
  i: 'italic',
  u: 'underline',
}

/** No `document`: o foco vive dentro do grid, um web component, e o evento nem sempre sobe ao React. */
export function useFormatShortcuts(onToggle: (key: 'bold' | 'italic' | 'underline') => void): void {
  // Montado uma vez: a grade renderiza a cada tecla.
  const toggle = useRef(onToggle)
  toggle.current = onToggle

  useEffect(() => {
    const shortcut = (event: KeyboardEvent): void => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey) return

      // Num campo de texto, o atalho é do campo.
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
