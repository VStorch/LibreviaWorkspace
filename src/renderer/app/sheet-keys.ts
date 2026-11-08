import { useEffect } from 'react'
import { SHORTCUTS, isShortcutPressed } from '@shared/shortcuts.js'
import { activeWorkspaceStore } from '../state/tabs.js'

const SHEET_KEYS = [
  { shortcut: SHORTCUTS.nextSheet, step: 1 },
  { shortcut: SHORTCUTS.previousSheet, step: -1 },
] as const

/** Stops at the first and the last sheet, as Excel does. */
function onKeyDown(event: KeyboardEvent): void {
  const pressed = SHEET_KEYS.find(({ shortcut }) => isShortcutPressed(shortcut, event))
  if (pressed === undefined) return
  const state = activeWorkspaceStore().getState()
  const { workbook } = state
  if (workbook === null) return
  event.preventDefault()
  event.stopPropagation()
  state.selectSheet(workbook.activeSheet + pressed.step)
}

/** `Ctrl+PgDn` and `Ctrl+PgUp` on a spreadsheet, heard before the grid. */
export function useSheetKeys(): void {
  useEffect(() => {
    window.addEventListener('keydown', onKeyDown, true)
    return () => window.removeEventListener('keydown', onKeyDown, true)
  }, [])
}
