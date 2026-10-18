import { useEffect, useRef } from 'react'

/** Set by a field with nothing typed yet, such as the formula bar: undo goes to the sheet. */
export const SHEET_UNDO_ATTRIBUTE = 'data-sheet-undo'

export const TEXT_FIELD = 'input, textarea, select, [contenteditable="true"]'
/** The app's dialogs float over the sheet without `aria-modal`; while one is open it has the keys. */
const DIALOG = '[role="dialog"], [aria-modal="true"], dialog[open]'

export const dialogOverSheet = (): boolean => document.querySelector(DIALOG) !== null

/** Excel's and LibreOffice's keys. Ctrl+Shift+Y is not one of them: it opens the assistant. */
function historyKeyOf(event: KeyboardEvent): 'undo' | 'redo' | null {
  if (!(event.ctrlKey || event.metaKey) || event.altKey) return null
  const key = event.key.toLowerCase()
  if (key === 'z') return event.shiftKey ? 'redo' : 'undo'
  if (key === 'y' && !event.shiftKey) return 'redo'
  return null
}

/**
 * The sheet's history, not the field's: outside any text field being typed into, and with no dialog
 * over the sheet.
 */
export function sheetOwnsUndo(target: EventTarget | null): boolean {
  if (dialogOverSheet()) return false
  if (!(target instanceof Element)) return true
  const field = target.closest(TEXT_FIELD)
  return field === null || field.hasAttribute(SHEET_UNDO_ATTRIBUTE)
}

/**
 * On `document`, as the format shortcuts: focus lives inside the grid. In a text field, as the cell
 * being typed into, the field undoes its own typing.
 */
export function useSheetUndo(onUndo: () => void, onRedo: () => void): void {
  const actions = useRef({ onUndo, onRedo })
  actions.current = { onUndo, onRedo }

  useEffect(() => {
    const shortcut = (event: KeyboardEvent): void => {
      const direction = historyKeyOf(event)
      if (direction === null || !sheetOwnsUndo(event.target)) return
      event.preventDefault()
      if (direction === 'redo') actions.current.onRedo()
      else actions.current.onUndo()
    }
    document.addEventListener('keydown', shortcut)
    return () => document.removeEventListener('keydown', shortcut)
  }, [])
}
