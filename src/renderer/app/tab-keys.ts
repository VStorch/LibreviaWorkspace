import { useEffect } from 'react'
import { SHORTCUTS, isShortcutPressed } from '@shared/shortcuts.js'
import { MenuCommand } from '@shared/types.js'
import { executeCommand } from '../commands/execute.js'
import { activeWorkspaceStore } from '../state/tabs.js'

type Step = 1 | -1

const PAGE_KEYS = [
  { shortcut: SHORTCUTS.nextTabByPage, step: 1, command: MenuCommand.NextTab },
  { shortcut: SHORTCUTS.previousTabByPage, step: -1, command: MenuCommand.PreviousTab },
] as const

/** On a spreadsheet the keys belong to its sheets, as in Excel, and stop at the first and the last. */
function stepSheet(step: Step): boolean {
  const state = activeWorkspaceStore().getState()
  const { workbook } = state
  if (workbook === null) return false
  const target = workbook.activeSheet + step
  if (target >= 0 && target < workbook.sheets.length) state.selectSheet(target)
  return true
}

function onKeyDown(event: KeyboardEvent): void {
  const pressed = PAGE_KEYS.find(({ shortcut }) => isShortcutPressed(shortcut, event))
  if (pressed === undefined) return
  event.preventDefault()
  event.stopPropagation()
  if (!stepSheet(pressed.step)) void executeCommand(pressed.command)
}

/** `Ctrl+PgDn` and `Ctrl+PgUp`, heard before the grid and the editor. */
export function useTabKeys(): void {
  useEffect(() => {
    window.addEventListener('keydown', onKeyDown, true)
    return () => window.removeEventListener('keydown', onKeyDown, true)
  }, [])
}
