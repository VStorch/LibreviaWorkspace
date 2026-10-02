import type { Editor } from '@tiptap/react'
import { TextSelection } from '@tiptap/pm/state'
import { useWorkspace } from '../state/workspace.js'
import { noteBodyOf } from './extensions/note-view.js'
import { SKIP_TRACKING } from './extensions/track-input.js'
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
 * estava — e fora do controle (`SKIP_TRACKING`): aceitar não é uma edição nova a
 * controlar. As regras moram em track-changes.ts; aqui só a ponte com o editor.
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
  editor.view.dispatch(tr.setMeta(SKIP_TRACKING, true).scrollIntoView())
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

/** Próxima e anterior: a alteração vira a seleção e entra na tela. Devolve se achou. */
export function goToChange(editor: Editor, direction: 1 | -1): boolean {
  const { selection } = editor.state
  const change = adjacentChange(editor.state.doc, direction === 1 ? selection.to : selection.from, direction)
  if (change === null) return false
  editor.view.focus()
  // Dentro de uma nota a seleção é a do corpo, que tem editor próprio: o texto
  // fica com o cursor depois da referência, e o corpo escolhe o trecho.
  const $from = editor.state.doc.resolve(change.from)
  for (let depth = $from.depth; depth > 0; depth--) {
    if ($from.node(depth).type.name !== 'noteRef') continue
    const reference = $from.before(depth)
    const tr = editor.state.tr.setSelection(TextSelection.create(editor.state.doc, $from.after(depth)))
    editor.view.dispatch(tr.scrollIntoView())
    const body = noteBodyOf(editor.view.nodeDOM(reference))
    if (body !== undefined) {
      body.view.focus()
      body.select(change.from - reference - 1, change.to - reference - 1)
    }
    return true
  }
  editor.view.dispatch(selectChange(editor.state.tr, change))
  return true
}
