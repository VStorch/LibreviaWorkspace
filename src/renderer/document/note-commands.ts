import type { Editor } from '@tiptap/react'
import { TextSelection } from '@tiptap/pm/state'
import { NoteKind } from '@services/document/notes.js'
import { noteBodyOf } from './extensions/note-view.js'
import { useWorkspace } from '../state/workspace.js'

/** O estilo que o Word dá ao parágrafo da nota nova, quando o documento o tem. */
const NOTE_STYLE: Record<NoteKind, string> = {
  [NoteKind.Footnote]: 'FootnoteText',
  [NoteKind.Endnote]: 'EndnoteText',
}

/**
 * Insere uma nota no cursor (M11): a referência no fim da seleção, como no Word,
 * e o cursor no corpo da nota, que começa com o espaço depois do número.
 *
 * O foco vai para o corpo quando ele chega à folha: antes de a paginação o pôr
 * no pé da página, ele mora num depósito escondido (ver `note-view.ts`).
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
