import { useEffect } from 'react'
import { parseClipboardText, pasteText, rangeAsText } from '@services/spreadsheet/clipboard.js'
import { clearContents, type Range } from '@services/spreadsheet/edit.js'
import type { Sheet } from '@services/spreadsheet/model.js'
import { TEXT_FIELD, dialogOverSheet } from './useSheetUndo.js'

const PLAIN_TEXT = 'text/plain'
const CLEAR_KEYS: ReadonlySet<string> = new Set(['Delete', 'Backspace'])

/** Outside a field and a dialog, and with no text of the page selected, as an answer of the assistant. */
function sheetOwns(event: Event, needsNoTextSelected: boolean): boolean {
  if (dialogOverSheet()) return false
  if (event.target instanceof Element && event.target.closest(TEXT_FIELD) !== null) return false
  return !needsNoTextSelected || document.getSelection()?.isCollapsed !== false
}

function isClearKey(event: KeyboardEvent): boolean {
  return CLEAR_KEYS.has(event.key) && !(event.ctrlKey || event.metaKey || event.altKey || event.shiftKey)
}

/**
 * On the model and the whole selection: the grid copies, pastes and clears only in the pane of its
 * focused cell, so with frozen columns it would leave out half the range, or paste nowhere. Copying
 * takes what the cells show, as the grid did.
 */
export function useSheetClipboard(
  current: { readonly current: Sheet },
  selected: { readonly current: Range },
  applyChange: (sheet: Sheet) => void,
): void {
  useEffect(() => {
    const copy = (event: ClipboardEvent): void => {
      if (event.clipboardData === null || !sheetOwns(event, true)) return
      event.preventDefault()
      event.stopImmediatePropagation()
      event.clipboardData.setData(PLAIN_TEXT, rangeAsText(current.current, selected.current))
      if (event.type === 'cut') applyChange(clearContents(current.current, selected.current))
    }

    const paste = (event: ClipboardEvent): void => {
      const text = event.clipboardData?.getData(PLAIN_TEXT)
      if (text === undefined || text === '' || !sheetOwns(event, false)) return
      event.preventDefault()
      event.stopImmediatePropagation()
      const { fromRow, fromColumn } = selected.current
      applyChange(pasteText(current.current, fromRow, fromColumn, parseClipboardText(text)))
    }

    const clear = (event: KeyboardEvent): void => {
      if (!isClearKey(event) || !sheetOwns(event, false)) return
      event.preventDefault()
      event.stopImmediatePropagation()
      applyChange(clearContents(current.current, selected.current))
    }

    // Before the grid's own listeners, on `document`.
    window.addEventListener('copy', copy, { capture: true })
    window.addEventListener('cut', copy, { capture: true })
    window.addEventListener('paste', paste, { capture: true })
    window.addEventListener('keydown', clear, { capture: true })
    return () => {
      window.removeEventListener('copy', copy, { capture: true })
      window.removeEventListener('cut', copy, { capture: true })
      window.removeEventListener('paste', paste, { capture: true })
      window.removeEventListener('keydown', clear, { capture: true })
    }
  }, [current, selected, applyChange])
}
