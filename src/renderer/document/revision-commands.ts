import type { Editor } from '@tiptap/react'
import { useWorkspace } from '../state/workspace.js'
import { caretOf, selectInNote } from './extensions/note-view.js'
import { SKIP_TRACKING } from './extensions/track-input.js'
import {
  adjacentChange,
  changeAt,
  selectChange,
  settleAllChanges,
  settleChangeAt,
} from './extensions/track-changes.js'

/**
 * Ordinary transactions, with undo, outside tracking (`SKIP_TRACKING`): accepting is not a new
 * edit.
 */

/** The cursor is the note's when the note is active (`caretOf`). */
export function hasChangeAtCursor(editor: Editor): boolean {
  return changeAt(editor.state.doc, caretOf(editor.view).from) !== null
}

export function settleChange(editor: Editor, accept: boolean): boolean {
  if (useWorkspace.getState().readOnly) return false
  const caret = caretOf(editor.view)
  const tr = editor.state.tr
  if (!settleChangeAt(tr, caret.from, accept)) return false
  tr.setMeta(SKIP_TRACKING, true)
  // In a note, the body gets the change through its own view, and focus returns to it.
  if (caret.note === null) tr.scrollIntoView()
  editor.view.dispatch(tr)
  caret.note?.view.focus()
  return true
}

export function settleAll(editor: Editor, accept: boolean): boolean {
  if (useWorkspace.getState().readOnly) return false
  const tr = editor.state.tr
  if (!settleAllChanges(tr, accept)) return false
  editor.view.dispatch(tr.setMeta(SKIP_TRACKING, true))
  return true
}

/** Returns whether one was found. From the active note's cursor, when there is one. */
export function goToChange(editor: Editor, direction: 1 | -1): boolean {
  const caret = caretOf(editor.view)
  const change = adjacentChange(editor.state.doc, direction === 1 ? caret.to : caret.from, direction)
  if (change === null) return false
  if (selectInNote(editor.view, change.from, change.to, true)) return true
  editor.view.focus()
  editor.view.dispatch(selectChange(editor.state.tr, change))
  return true
}
