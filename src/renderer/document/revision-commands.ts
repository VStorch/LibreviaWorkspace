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

/** Transações comuns, com desfazer, fora do controle (`SKIP_TRACKING`): aceitar não é edição nova. */

/** O cursor é o da nota quando ela é a ativa (`caretOf`). */
export function hasChangeAtCursor(editor: Editor): boolean {
  return changeAt(editor.state.doc, caretOf(editor.view).from) !== null
}

export function settleChange(editor: Editor, accept: boolean): boolean {
  if (useWorkspace.getState().readOnly) return false
  const caret = caretOf(editor.view)
  const tr = editor.state.tr
  if (!settleChangeAt(tr, caret.from, accept)) return false
  tr.setMeta(SKIP_TRACKING, true)
  // Na nota, o corpo recebe a mudança pela própria vista, e o foco volta a ele.
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

/** Devolve se achou. A partir do cursor da nota ativa, quando há uma. */
export function goToChange(editor: Editor, direction: 1 | -1): boolean {
  const caret = caretOf(editor.view)
  const change = adjacentChange(editor.state.doc, direction === 1 ? caret.to : caret.from, direction)
  if (change === null) return false
  if (selectInNote(editor.view, change.from, change.to, true)) return true
  editor.view.focus()
  editor.view.dispatch(selectChange(editor.state.tr, change))
  return true
}
