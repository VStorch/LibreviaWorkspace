import type { Editor } from '@tiptap/react'
import { useWorkspace } from '../state/workspace.js'
import {
  adjacentChange,
  changeAt,
  selectChange,
  settleAllChanges,
  settleChangeAt,
} from './extensions/track-changes.js'

/**
 * Aceitar, rejeitar e andar entre as alterações (M10, controle de alterações).
 *
 * Cada comando é uma transação comum do editor: o desfazer devolve a revisão como
 * estava. As regras moram em track-changes.ts; aqui só a ponte com o editor.
 */

/** Há uma alteração no cursor — é o que o menu de contexto pergunta. */
export function hasChangeAtCursor(editor: Editor): boolean {
  return changeAt(editor.state.doc, editor.state.selection.from) !== null
}

/** Aceita (ou rejeita) a alteração no cursor, e passa à seguinte, como no Word. */
export function settleChange(editor: Editor, accept: boolean): boolean {
  if (useWorkspace.getState().readOnly) return false
  const tr = editor.state.tr
  if (!settleChangeAt(tr, editor.state.selection.from, accept)) return false
  editor.view.dispatch(tr.scrollIntoView())
  return true
}

/** Aceita (ou rejeita) todas as alterações do documento. */
export function settleAll(editor: Editor, accept: boolean): boolean {
  if (useWorkspace.getState().readOnly) return false
  const tr = editor.state.tr
  if (!settleAllChanges(tr, accept)) return false
  editor.view.dispatch(tr)
  return true
}

/** Próxima e anterior: a alteração vira a seleção e entra na tela. Devolve se achou. */
export function goToChange(editor: Editor, direction: 1 | -1): boolean {
  const { selection } = editor.state
  const change = adjacentChange(editor.state.doc, direction === 1 ? selection.to : selection.from, direction)
  if (change === null) return false
  editor.view.focus()
  editor.view.dispatch(selectChange(editor.state.tr, change))
  return true
}
