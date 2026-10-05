import type { Editor } from '@tiptap/react'
import type { DocumentComment } from '@services/document/model.js'
import { initialsOf, nextCommentId, paragraphsOfText, resolveComments } from '@services/document/comments.js'
import { t } from '../i18n.js'
import { currentPreferences, setPreference } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import {
  adjacentComment,
  commentAnchorsOf,
  commentSelectionOf,
  commentsKey,
  insertCommentAnchors,
  removeCommentAnchors,
  selectComment,
} from './extensions/comment.js'
import { caretOf, selectInNote } from './extensions/note-view.js'

/**
 * Like sections: the ends in the text and the body in the store's library. Insert and delete are
 * one editor transaction each, with undo; text, reply and resolved only change the store, like
 * styles.
 */

/** A draft older than comments (`.sdoc` < 7) cannot store them: the command refuses. */
export function commentEditsAllowed(): boolean {
  const store = useWorkspace.getState()
  if (store.readOnly) return false
  if (!store.beforeComments) return true
  store.showError({ code: 'INTERNAL', message: t('comments.legacyDraft') })
  return false
}

function newComment(library: readonly DocumentComment[], parentId?: string): DocumentComment {
  const author = currentPreferences().authorName.trim()
  const initials = initialsOf(author)
  return {
    id: nextCommentId(library),
    author,
    ...(initials === '' ? {} : { initials }),
    date: new Date().toISOString(),
    paragraphs: [],
    done: false,
    ...(parentId === undefined ? {} : { parentId }),
  }
}

/**
 * A new card with the box open; empty, it gives up and takes the ends with it (`cancelNewComment`).
 */
export function insertComment(editor: Editor): void {
  if (!commentEditsAllowed()) return
  const store = useWorkspace.getState()
  const wasClean = !store.isDirty
  const comment = newComment(store.comments)
  // The library before the text, as with sections. No `focus()`: Tiptap's focus would arrive later
  // and steal the card box's.
  const caret = caretOf(editor.view)
  // A new comment inside a note is refused: LibreOffice does not open a .docx with
  // `w:commentReference` in `footnotes.xml`.
  if (caret.note !== null) {
    store.showError({ code: 'INTERNAL', message: t('comments.notInNote') })
    return
  }
  const inserted = editor
    .chain()
    .command(({ tr, state }) => {
      if (!insertCommentAnchors(tr, state.schema, comment.id, caret)) return false
      store.setComments([...store.comments, comment])
      return true
    })
    .run()
  if (!inserted) return
  showCommentsPane()
  cleanBeforeDraft = wasClean
  store.setCommentDraft(comment.id)
}

/** Giving up the new card brings the document back to saved. */
let cleanBeforeDraft = false

/** Outside undo: redoing a comment without text would be an empty card. */
export function cancelNewComment(editor: Editor, cid: string): void {
  const store = useWorkspace.getState()
  store.setCommentDraft(null)
  const tr = editor.state.tr
  if (removeCommentAnchors(tr, cid)) editor.view.dispatch(tr.setMeta('addToHistory', false))
  if (cleanBeforeDraft) useWorkspace.setState({ isDirty: false })
  cleanBeforeDraft = false
}

/** The ends go, and with them the card and the replies. */
export function deleteCommentThread(editor: Editor, cid: string): void {
  if (!commentEditsAllowed()) return
  const tr = editor.state.tr
  if (removeCommentAnchors(tr, cid)) editor.view.dispatch(tr)
  const store = useWorkspace.getState()
  if (store.commentDraft === cid) store.setCommentDraft(null)
}

export function editComment(id: string, text: string): void {
  if (!commentEditsAllowed()) return
  const store = useWorkspace.getState()
  const paragraphs = paragraphsOfText(text)
  store.setComments(
    store.comments.map((comment) => (comment.id === id ? { ...comment, paragraphs } : comment)),
  )
  if (store.commentDraft === id) store.setCommentDraft(null)
}

export function replyToComment(rootId: string, text: string): void {
  if (!commentEditsAllowed()) return
  const store = useWorkspace.getState()
  const reply = { ...newComment(store.comments, rootId), paragraphs: paragraphsOfText(text) }
  store.setComments([...store.comments, reply])
}

/** Applies to the whole thread. */
export function setCommentDone(rootId: string, done: boolean): void {
  if (!commentEditsAllowed()) return
  const store = useWorkspace.getState()
  store.setComments(store.comments.map((comment) => (comment.id === rootId ? { ...comment, done } : comment)))
}

function showCommentsPane(): void {
  if (!currentPreferences().commentsPane) void setPreference({ commentsPane: true })
}

/** The thread's card stays selected. Returns whether it found any. */
export function goToComment(editor: Editor, direction: 1 | -1): boolean {
  const store = useWorkspace.getState()
  const anchors = commentAnchorsOf(editor.state.doc)
  const visible = resolveComments(new Set(anchors.keys()), store.comments, new Set(store.commentsOutside))
  const ids = new Set(visible.map((comment) => comment.id))
  const threads = new Set(
    visible
      .filter((comment) => comment.parentId === undefined || !ids.has(comment.parentId))
      .map((comment) => comment.id),
  )
  const active = commentsKey.getState(editor.state)?.active ?? null
  const cid = adjacentComment(editor.state.doc, threads, active, caretOf(editor.view).from, direction)
  if (cid === null) return false
  showCommentsPane()
  showComment(editor, cid, true)
  return true
}

/** In the note body, when that is where it is. `focus` takes the keyboard to the range. */
export function showComment(editor: Editor, cid: string, focus: boolean): void {
  const anchor = commentAnchorsOf(editor.state.doc).get(cid)
  const range = anchor === undefined ? null : commentSelectionOf(anchor)
  if (focus) editor.view.focus()
  editor.view.dispatch(selectComment(editor.state.tr, cid))
  if (range !== null) selectInNote(editor.view, range.from, range.to, focus)
}
