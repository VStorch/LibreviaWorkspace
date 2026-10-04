import { useEffect, type RefObject } from 'react'
import type { Editor } from '@tiptap/react'
import { RevisionView, type EditorPreferences } from '@shared/types.js'
import type { StyleSheet } from '@services/document/styles.js'
import { revisionViewOf, setRevisionViewMeta } from './extensions/revision-view.js'
import { noteBodiesOf, setNotePool } from './extensions/note-view.js'
import { focusComment } from './extensions/comment.js'

export interface EditorSyncState {
  readonly readOnly: boolean
  readonly reading: boolean
  readonly revisionView: RevisionView
  readonly styles: StyleSheet
  readonly preferences: EditorPreferences
  readonly notePool: RefObject<HTMLDivElement | null>
  /** A paginação mede de novo. */
  readonly touch: () => void
}

/** O editor já montado acompanha a loja e as preferências, sem ser recriado. */
export function useEditorSync(editor: Editor | null, state: EditorSyncState): void {
  const { readOnly, reading, revisionView, styles, preferences, notePool, touch } = state

  // `setEditable` no editor já montado, e com o segundo argumento: recriar o
  // editor perderia cursor e histórico, e o update padrão marcaria todo arquivo
  // aberto como "não salvo". O modo de leitura e o Original também travam a edição.
  const original = revisionView === RevisionView.Original
  useEffect(() => {
    editor?.setEditable(!readOnly && !reading && !original, false)
    if (editor !== null && !editor.isDestroyed) for (const body of noteBodiesOf(editor.view)) body.refresh()
  }, [editor, readOnly, reading, original])

  useEffect(() => {
    if (editor === null || editor.isDestroyed) return undefined
    const view = editor.view
    setNotePool(view, notePool.current)
    return () => setNotePool(view, null)
  }, [editor, notePool])

  // Uma transação sem mudança no documento: a paginação mede de novo, e o escondido não ocupa lugar.
  useEffect(() => {
    if (editor === null || editor.isDestroyed || revisionViewOf(editor.state) === revisionView) return
    editor.view.dispatch(setRevisionViewMeta(editor.state.tr, revisionView))
    // O CSS esconde pela classe do modo (`revisions-…`), que o corpo das notas também leva.
    for (const body of noteBodiesOf(editor.view)) body.refresh()
    touch()
  }, [editor, revisionView, touch])

  useEffect(() => {
    if (editor !== null) editor.storage.paragraphCommands.styles = styles
  }, [editor, styles])

  /**
   * Escrito no elemento: o ProseMirror só lê os atributos ao criar a visão. O
   * main liga o corretor na sessão; isto é a outra metade.
   */
  useEffect(() => {
    editor?.view.dom.setAttribute('spellcheck', preferences.spellcheck ? 'true' : 'false')
  }, [editor, preferences.spellcheck])

  // A transação das marcas de formatação não muda o documento.
  useEffect(() => {
    editor?.commands.showInvisibleCharacters(preferences.invisibleCharacters)
  }, [editor, preferences.invisibleCharacters])

  // Sem o painel, sai só o realce: pontas e corpos ficam, e voltam ao arquivo.
  useEffect(() => {
    if (editor === null || editor.isDestroyed) return
    editor.view.dispatch(focusComment(editor.state.tr, { hidden: !preferences.commentsPane }))
  }, [editor, preferences.commentsPane])
}
