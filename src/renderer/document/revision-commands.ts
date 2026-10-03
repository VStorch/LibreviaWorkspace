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
 * Aceitar, rejeitar e andar entre as alterações.
 *
 * Cada comando é uma transação comum do editor: o desfazer devolve a revisão como
 * estava — e fora do controle (`SKIP_TRACKING`): aceitar não é uma edição nova a
 * controlar. As regras moram em track-changes.ts; aqui só a ponte com o editor.
 */

/**
 * Há uma alteração no cursor — é o que o menu de contexto pergunta. O cursor é o
 * da nota quando ela é a ativa (`caretOf`): o corpo tem editor próprio.
 */
export function hasChangeAtCursor(editor: Editor): boolean {
  return changeAt(editor.state.doc, caretOf(editor.view).from) !== null
}

/** Aceita (ou rejeita) a alteração no cursor — no texto ou na nota em que ele está. */
export function settleChange(editor: Editor, accept: boolean): boolean {
  if (useWorkspace.getState().readOnly) return false
  const caret = caretOf(editor.view)
  const tr = editor.state.tr
  if (!settleChangeAt(tr, caret.from, accept)) return false
  tr.setMeta(SKIP_TRACKING, true)
  // Na nota, a seleção do texto fica onde está (depois da referência), e o
  // corpo recebe a mudança pela própria vista; o foco volta para ele, que o
  // clique no menu de contexto tirou.
  if (caret.note === null) tr.scrollIntoView()
  editor.view.dispatch(tr)
  caret.note?.view.focus()
  return true
}

/** Aceita (ou rejeita) todas as alterações do documento. */
export function settleAll(editor: Editor, accept: boolean): boolean {
  if (useWorkspace.getState().readOnly) return false
  const tr = editor.state.tr
  if (!settleAllChanges(tr, accept)) return false
  editor.view.dispatch(tr.setMeta(SKIP_TRACKING, true))
  return true
}

/**
 * Próxima e anterior: a alteração vira a seleção e entra na tela. Devolve se
 * achou. A partir do cursor da nota, quando ela é a ativa: as alterações seguintes
 * da mesma nota vêm antes do texto depois da referência.
 */
export function goToChange(editor: Editor, direction: 1 | -1): boolean {
  const caret = caretOf(editor.view)
  const change = adjacentChange(editor.state.doc, direction === 1 ? caret.to : caret.from, direction)
  if (change === null) return false
  // Dentro de uma nota a seleção é a do corpo, que tem editor próprio.
  if (selectInNote(editor.view, change.from, change.to, true)) return true
  editor.view.focus()
  editor.view.dispatch(selectChange(editor.state.tr, change))
  return true
}
