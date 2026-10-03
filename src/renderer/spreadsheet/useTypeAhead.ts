import { useCallback, useEffect, useMemo, useRef } from 'react'

/**
 * Depois do Enter o grid espera 70 ms (`RESIZE_INTERVAL + 30`) antes de mover o
 * foco: quem digita sem parar acerta essa janela, e `1200` abaixo de `980` virava
 * `200`. A tecla é guardada e devolvida quando o foco chega, pelo
 * `beforekeydown` que a própria biblioteca oferece.
 */
export interface TypeAhead {
  readonly begin: () => void
  readonly settle: () => void
}

/** Para o commit que não move o foco, como confirmar clicando noutra célula. */
const WINDOW_MS = 250

export function useTypeAhead(readOnly: boolean): TypeAhead {
  const typed = useRef<string[]>([])
  /** Cada sobreposição de seleção do grid emite o mesmo `beforekeydown`: sem isto a tecla seria guardada várias vezes. */
  const lastHeld = useRef<KeyboardEvent | null>(null)
  const open = useRef(false)
  const timer = useRef<number | null>(null)

  const closeWindow = useCallback(() => {
    open.current = false
    if (timer.current !== null) {
      window.clearTimeout(timer.current)
      timer.current = null
    }
  }, [])

  /** Pelo `keydown` do grid, e não escrevendo na célula: uma definição só de "digitar por cima". */
  const replay = useCallback(() => {
    const held = typed.current
    typed.current = []
    if (readOnly) return

    for (const key of held) {
      document.dispatchEvent(
        new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, composed: true }),
      )
    }
  }, [readOnly])

  // Indireção: trocar o temporizador a cada renderização o reiniciaria no meio da janela.
  const replayRef = useRef(replay)
  replayRef.current = replay

  const begin = useCallback(() => {
    open.current = true
    if (timer.current !== null) window.clearTimeout(timer.current)
    timer.current = window.setTimeout(() => {
      open.current = false
      timer.current = null
      replayRef.current()
    }, WINDOW_MS)
  }, [])

  const settle = useCallback(() => {
    closeWindow()
    replayRef.current()
  }, [closeWindow])

  useEffect(() => {
    /** Já no Enter de um editor de célula (`.edit-input-wrapper`, como `isEditInput`). */
    const commit = (event: KeyboardEvent): void => {
      if (readOnly || !event.isTrusted) return
      if (event.key !== 'Enter' && event.key !== 'Tab') return
      if (!(event.target instanceof HTMLElement)) return
      if (event.target.closest('.edit-input-wrapper') === null) return
      begin()
    }

    // No `document`, onde o grid escuta o `keydown`, na mesma pilha da resposta.
    const hold = (event: Event): void => {
      if (!open.current || readOnly) return

      const original = (event as CustomEvent<{ original: KeyboardEvent }>).detail.original
      if (!original.isTrusted) return
      if (original === lastHeld.current) {
        event.preventDefault()
        return
      }
      if (original.ctrlKey || original.metaKey || original.altKey) return
      // Só caractere digitável: as setas continuam navegando.
      if (original.key.length !== 1) return

      // Senão o grid a trataria na célula anterior: `1200` virava `9801200`.
      event.preventDefault()
      original.preventDefault()
      lastHeld.current = original
      typed.current.push(original.key)
    }

    document.addEventListener('keydown', commit, true)
    document.addEventListener('beforekeydown', hold)
    return () => {
      document.removeEventListener('keydown', commit, true)
      document.removeEventListener('beforekeydown', hold)
    }
  }, [begin, readOnly])

  return useMemo(() => ({ begin, settle }), [begin, settle])
}
