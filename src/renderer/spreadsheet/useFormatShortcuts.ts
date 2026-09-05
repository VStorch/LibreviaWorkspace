import { useEffect, useRef } from 'react'

/** O atributo que cada atalho alterna. */
const SHORTCUTS: Record<string, 'bold' | 'italic' | 'underline' | undefined> = {
  b: 'bold',
  i: 'italic',
  u: 'underline',
}

/**
 * Ctrl+B, Ctrl+I e Ctrl+U na grade.
 *
 * No documento é o TipTap que trata; aqui não há editor de texto, então o atalho
 * é nosso. Fica no `document` porque o foco vive dentro do grid, que é um web
 * component — o evento nem sempre sobe até um `onKeyDown` do React.
 */
export function useFormatShortcuts(onToggle: (key: 'bold' | 'italic' | 'underline') => void): void {
  // O ouvinte é montado uma vez só: guardar a ação numa referência evita
  // reassiná-lo a cada renderização, e a grade renderiza a cada tecla.
  const toggle = useRef(onToggle)
  toggle.current = onToggle

  useEffect(() => {
    const shortcut = (event: KeyboardEvent): void => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey) return

      // Digitando na barra de fórmulas ou no editor da célula, o atalho é do
      // campo de texto: formatar a célula por baixo seria o oposto do esperado.
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
