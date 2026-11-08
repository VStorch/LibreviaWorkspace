import { useCallback, useEffect } from 'react'
import { parseClipboardText, pasteText, rangeAsText } from '@services/spreadsheet/clipboard.js'
import { clearContents, type Range } from '@services/spreadsheet/edit.js'
import type { Sheet } from '@services/spreadsheet/model.js'
import { SHORTCUTS, isShortcutPressed } from '@shared/shortcuts.js'
import { useEditorHandlers } from '../commands/useEditorHandlers.js'
import { EditorCommand } from '../document/editor-commands.js'
import { useWorkspace } from '../state/workspace.js'
import { TEXT_FIELD, dialogOverSheet } from './useSheetUndo.js'

const PLAIN_TEXT = 'text/plain'
const PASTE_COMMANDS = [EditorCommand.PasteWithoutFormat]
const CLEAR_KEYS: ReadonlySet<string> = new Set(['Delete', 'Backspace'])

/** Outside a field and a dialog, and with no text of the page selected, as an answer of the assistant. */
function sheetOwns(target: EventTarget | null, needsNoTextSelected: boolean): boolean {
  if (dialogOverSheet()) return false
  if (target instanceof Element && target.closest(TEXT_FIELD) !== null) return false
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
      if (event.clipboardData === null || !sheetOwns(event.target, true)) return
      event.preventDefault()
      event.stopImmediatePropagation()
      event.clipboardData.setData(PLAIN_TEXT, rangeAsText(current.current, selected.current))
      if (event.type === 'cut') applyChange(clearContents(current.current, selected.current))
    }

    const paste = (event: ClipboardEvent): void => {
      const text = event.clipboardData?.getData(PLAIN_TEXT)
      if (text === undefined || text === '' || !sheetOwns(event.target, false)) return
      event.preventDefault()
      event.stopImmediatePropagation()
      const { fromRow, fromColumn } = selected.current
      applyChange(pasteText(current.current, fromRow, fromColumn, parseClipboardText(text)))
    }

    const clear = (event: KeyboardEvent): void => {
      if (!isClearKey(event) || !sheetOwns(event.target, false)) return
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

/**
 * Pasting already drops the formatting, so this is the same paste, from the menu and from
 * `Ctrl+Shift+V`, which Chromium would otherwise turn into its own paste.
 */
export function useSheetPasteWithoutFormat(
  current: { readonly current: Sheet },
  selected: { readonly current: Range },
  applyChange: (sheet: Sheet) => void,
): void {
  const showError = useWorkspace((state) => state.showError)

  const paste = useCallback(async (): Promise<void> => {
    const result = await window.api.edit.readClipboardText({})
    if (!result.ok) {
      showError(result.error)
      return
    }
    if (result.data.text === '') return
    const { fromRow, fromColumn } = selected.current
    applyChange(pasteText(current.current, fromRow, fromColumn, parseClipboardText(result.data.text)))
  }, [current, selected, applyChange, showError])

  useEffect(() => {
    const key = (event: KeyboardEvent): void => {
      if (!isShortcutPressed(SHORTCUTS.pasteWithoutFormat, event) || !sheetOwns(event.target, false)) return
      event.preventDefault()
      event.stopImmediatePropagation()
      void paste()
    }
    window.addEventListener('keydown', key, { capture: true })
    return () => window.removeEventListener('keydown', key, { capture: true })
  }, [paste])

  useEditorHandlers(PASTE_COMMANDS, () => {
    if (sheetOwns(document.activeElement, false)) void paste()
  })
}
