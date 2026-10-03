import { useEffect } from 'react'
import { setPreference, usePreferences } from './preferences.js'

/** A preferência mora no main, que marca o item do menu "Exibir". */
export function useReadingMode(): boolean {
  return usePreferences((state) => state.preferences.readingMode)
}

export async function setReadingMode(readingMode: boolean): Promise<void> {
  await setPreference({ readingMode })
}

/**
 * Sem barras na tela, Esc é o que impede a pessoa de se sentir trancada. Só com
 * o modo ligado: fora dele, Esc é do painel ou diálogo aberto.
 */
export function useLeaveReadingOnEscape(active: boolean): void {
  useEffect(() => {
    if (!active) return undefined

    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key !== 'Escape' || event.defaultPrevented) return
      event.preventDefault()
      void setReadingMode(false)
    }

    // Na captura: o ProseMirror ainda recebe as teclas e consumiria o Esc.
    window.addEventListener('keydown', onKeyDown, true)
    return () => window.removeEventListener('keydown', onKeyDown, true)
  }, [active])
}
