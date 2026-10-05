import type { Editor } from '@tiptap/react'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { NoteKind } from '@services/document/notes.js'
import { activeNoteOf, noteBodyOf } from './extensions/note-view.js'
import { useWorkspace } from '../state/workspace.js'

/** When the document has it. */
const NOTE_STYLE: Record<NoteKind, string> = {
  [NoteKind.Footnote]: 'FootnoteText',
  [NoteKind.Endnote]: 'EndnoteText',
}

/**
 * The reference at the end of the selection, as in Word, and the cursor in the body. Focus goes to
 * the body when it reaches the sheet (see `note-view.ts`).
 */
export function insertNote(editor: Editor, kind: NoteKind): boolean {
  const { state, schema } = editor
  const type = schema.nodes['noteRef']
  const paragraph = schema.nodes['paragraph']
  if (type === undefined || paragraph === undefined) return false
  const { $to } = state.selection
  if (!$to.parent.inlineContent || !$to.parent.type.contentMatch.matchType(type)) return false

  const style = NOTE_STYLE[kind]
  const attrs = useWorkspace.getState().styles.styles[style] === undefined ? {} : { styleId: style }
  const note = type.create({ kind }, paragraph.create(attrs, schema.text(' ')))
  const at = $to.pos
  const tr = state.tr.insert(at, note)
  tr.setSelection(TextSelection.create(tr.doc, at + note.nodeSize)).scrollIntoView()
  editor.view.dispatch(tr)
  const body = noteBodyOf(editor.view.nodeDOM(at))
  if (body !== undefined) {
    body.separateHistory = true
    body.reveal()
  }
  return true
}

/** The body holding the cursor, the selected reference, or the one touching the cursor. */
export function noteAtCursor(editor: Editor): { pos: number; kind: NoteKind } | null {
  const { state, view } = editor
  const found = (pos: number | undefined): { pos: number; kind: NoteKind } | null => {
    const node = pos === undefined ? null : state.doc.nodeAt(pos)
    if (pos === undefined || node === null || node.type.name !== 'noteRef') return null
    return { pos, kind: node.attrs['kind'] === NoteKind.Endnote ? NoteKind.Endnote : NoteKind.Footnote }
  }
  const active = activeNoteOf(view)
  if (active !== null) return found(active.position())
  const { selection } = state
  if (selection instanceof NodeSelection) return found(selection.from)
  if (!selection.empty) return null
  const { $from } = selection
  if ($from.nodeBefore?.type.name === 'noteRef') return found($from.pos - $from.nodeBefore.nodeSize)
  if ($from.nodeAfter?.type.name === 'noteRef') return found($from.pos)
  return null
}

/** The reference loses its `nid`: saving creates the note in the other part, with the same body. */
export function convertNote(editor: Editor, pos: number): boolean {
  if (useWorkspace.getState().readOnly) return false
  const { state } = editor
  const node = state.doc.nodeAt(pos)
  if (node === null || node.type.name !== 'noteRef') return false
  const from: NoteKind = node.attrs['kind'] === NoteKind.Endnote ? NoteKind.Endnote : NoteKind.Footnote
  const to: NoteKind = from === NoteKind.Endnote ? NoteKind.Footnote : NoteKind.Endnote
  const known = useWorkspace.getState().styles.styles[NOTE_STYLE[to]] !== undefined
  const tr = state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, kind: to, nid: null })
  node.descendants((child, offset) => {
    if (child.attrs['styleId'] !== NOTE_STYLE[from]) return true
    tr.setNodeMarkup(pos + 1 + offset, undefined, { ...child.attrs, styleId: known ? NOTE_STYLE[to] : null })
    return false
  })
  editor.view.dispatch(tr)
  return true
}
