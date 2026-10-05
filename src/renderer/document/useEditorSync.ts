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
  /** Pagination measures again. */
  readonly touch: () => void
}

/** The mounted editor follows the store and the preferences without being recreated. */
export function useEditorSync(editor: Editor | null, state: EditorSyncState): void {
  const { readOnly, reading, revisionView, styles, preferences, notePool, touch } = state

  // `setEditable` on the mounted editor, with the second argument: recreating the editor would lose
  // the cursor and the history, and the default update would mark every opened file as unsaved.
  // Reading mode and Original also lock editing.
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

  // A transaction without document changes: pagination measures again, and hidden text takes no
  // space.
  useEffect(() => {
    if (editor === null || editor.isDestroyed || revisionViewOf(editor.state) === revisionView) return
    editor.view.dispatch(setRevisionViewMeta(editor.state.tr, revisionView))
    // CSS hides by the mode class (`revisions-…`), which note bodies also carry.
    for (const body of noteBodiesOf(editor.view)) body.refresh()
    touch()
  }, [editor, revisionView, touch])

  useEffect(() => {
    if (editor !== null) editor.storage.paragraphCommands.styles = styles
  }, [editor, styles])

  /**
   * Written on the element: ProseMirror only reads the attributes when creating the view. Main
   * turns the spellchecker on in the session; this is the other half.
   */
  useEffect(() => {
    editor?.view.dom.setAttribute('spellcheck', preferences.spellcheck ? 'true' : 'false')
  }, [editor, preferences.spellcheck])

  // The formatting marks transaction does not change the document.
  useEffect(() => {
    editor?.commands.showInvisibleCharacters(preferences.invisibleCharacters)
  }, [editor, preferences.invisibleCharacters])

  // Without the pane, only the highlight goes: ends and bodies stay, and go back to the file.
  useEffect(() => {
    if (editor === null || editor.isDestroyed) return
    editor.view.dispatch(focusComment(editor.state.tr, { hidden: !preferences.commentsPane }))
  }, [editor, preferences.commentsPane])
}
