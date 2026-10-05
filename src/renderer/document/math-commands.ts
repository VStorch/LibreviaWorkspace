import type { Editor } from '@tiptap/react'
import { NodeSelection, type EditorState } from '@tiptap/pm/state'
import { DELETION, INSERTION } from './extensions/track-changes.js'

/**
 * One transaction, one undo step. Replacing is `replaceWith`: with tracking on, `track-input.ts`
 * sees a deletion and an insertion, as Word records it.
 */

export type EquationTarget =
  { readonly kind: 'insert'; readonly display: boolean } | { readonly kind: 'edit'; readonly pos: number }

/** `omml` goes null: the sidecar rebuilds it from the MathML. */
export interface EquationContent {
  readonly latex: string
  readonly mathml: string
  readonly display: boolean
}

export function equationAtSelection(state: EditorState): number | null {
  const selection = state.selection
  return selection instanceof NodeSelection && selection.node.type.name === 'math' ? selection.from : null
}

function attrsOf(content: EquationContent): Record<string, unknown> {
  return {
    omml: null,
    mathml: content.mathml,
    latex: content.latex,
    display: content.display,
    jc: null,
    lossy: [],
    editable: true,
  }
}

export function insertEquation(editor: Editor, content: EquationContent): void {
  editor
    .chain()
    .focus()
    .insertContent({ type: 'math', attrs: attrsOf(content) })
    .run()
}

/** False when there is no longer an equation there. */
export function replaceEquation(editor: Editor, pos: number, content: EquationContent): boolean {
  const node = editor.state.doc.nodeAt(pos)
  if (node === null || node.type.name !== 'math') return false

  const jc = content.display === node.attrs['display'] ? (node.attrs['jc'] as string | null) : null
  // The revision marks belong to the old equation.
  const marks = node.marks.filter((mark) => mark.type.name !== INSERTION && mark.type.name !== DELETION)
  const replacement = node.type.create({ ...attrsOf(content), jc }, null, marks)
  return editor
    .chain()
    .focus()
    .command(({ tr }) => {
      tr.replaceWith(pos, pos + node.nodeSize, replacement)
      return true
    })
    .run()
}
