import { useCallback, useEffect, useMemo, useRef } from 'react'

/**
 * After Enter the grid waits 70 ms (`RESIZE_INTERVAL + 30`) before moving focus: someone typing
 * nonstop hits that window, and `1200` below `980` became `200`. The key is held and replayed when
 * focus arrives, through the library's own `beforekeydown`.
 */
export interface TypeAhead {
  readonly begin: () => void
  readonly settle: () => void
}

/** For a commit that does not move focus, such as confirming by clicking another cell. */
const WINDOW_MS = 250

export function useTypeAhead(readOnly: boolean): TypeAhead {
  const typed = useRef<string[]>([])
  /**
   * Each grid selection overlay emits the same `beforekeydown`: without this the key would be held
   * several times.
   */
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

  /**
   * Through the grid's `keydown`, not by writing into the cell: a single definition of "type over".
   */
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

  // Indirection: replacing the timer on every render would restart it mid-window.
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
    /** Already on Enter in a cell editor (`.edit-input-wrapper`, like `isEditInput`). */
    const commit = (event: KeyboardEvent): void => {
      if (readOnly || !event.isTrusted) return
      if (event.key !== 'Enter' && event.key !== 'Tab') return
      if (!(event.target instanceof HTMLElement)) return
      if (event.target.closest('.edit-input-wrapper') === null) return
      begin()
    }

    // On `document`, where the grid listens for `keydown`, in the same stack as the response.
    const hold = (event: Event): void => {
      if (!open.current || readOnly) return

      const original = (event as CustomEvent<{ original: KeyboardEvent }>).detail.original
      if (!original.isTrusted) return
      if (original === lastHeld.current) {
        event.preventDefault()
        return
      }
      if (original.ctrlKey || original.metaKey || original.altKey) return
      // Only typeable characters: the arrow keys keep navigating.
      if (original.key.length !== 1) return

      // Otherwise the grid would handle it in the previous cell: `1200` became `9801200`.
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
