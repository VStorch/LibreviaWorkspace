import type { Editor } from '@tiptap/react'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { NoteKind } from '@services/document/notes.js'
import { activeNoteOf, noteBodyOf } from './extensions/note-view.js'
import { useWorkspace } from '../state/workspace.js'

/** O estilo que o Word dá ao parágrafo da nota nova, quando o documento o tem. */
const NOTE_STYLE: Record<NoteKind, string> = {
  [NoteKind.Footnote]: 'FootnoteText',
  [NoteKind.Endnote]: 'EndnoteText',
}

/**
 * Insere uma nota no cursor: a referência no fim da seleção, como no Word, e o
 * cursor no corpo da nota, que começa com o espaço depois do número.
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

/**
 * A nota sobre a qual o menu de contexto age: a do corpo em que está o cursor, a
 * referência selecionada, ou a que encosta no cursor do texto. `null` longe de
 * qualquer nota.
 */
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

/**
 * Converte a nota de rodapé em nota de fim, ou o contrário. A referência perde
 * o `nid`: a nota sai de uma parte do arquivo e a gravação a cria na outra, com
 * o mesmo corpo. O parágrafo com o estilo de um tipo passa ao do outro.
 */
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
