import { useEffect } from 'react'
import { setPreference, usePreferences } from './preferences.js'

/** The preference lives in main, which checks the "View" menu item. */
export function useReadingMode(): boolean {
  return usePreferences((state) => state.preferences.readingMode)
}

export async function setReadingMode(readingMode: boolean): Promise<void> {
  await setPreference({ readingMode })
}

/**
 * Without toolbars on screen, Esc is what keeps the user from feeling locked in. Only with the mode
 * on: outside it, Esc belongs to the open panel or dialog.
 */
export function useLeaveReadingOnEscape(active: boolean): void {
  useEffect(() => {
    if (!active) return undefined

    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key !== 'Escape' || event.defaultPrevented) return
      event.preventDefault()
      void setReadingMode(false)
    }

    // In the capture phase: ProseMirror still receives keys and would consume Esc.
    window.addEventListener('keydown', onKeyDown, true)
    return () => window.removeEventListener('keydown', onKeyDown, true)
  }, [active])
}
