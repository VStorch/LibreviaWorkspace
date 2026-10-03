import type { Editor } from '@tiptap/react'
import { NodeSelection, type EditorState } from '@tiptap/pm/state'
import { DELETION, INSERTION } from './extensions/track-changes.js'

/**
 * Inserir e trocar equação — o que o editor de equações faz ao OK.
 *
 * As duas são uma transação só, e por isso um passo só de desfazer. Trocar é
 * substituir o nó inteiro (`replaceWith`), e não mexer nos atributos: com o
 * controle de alterações ligado, `track-input.ts` vê a substituição como a
 * exclusão da equação antiga e a inserção da nova, que é como o Word a registra.
 */

/** O que o editor de equações abriu: uma equação nova, ou a do documento em `pos`. */
export type EquationTarget =
  { readonly kind: 'insert'; readonly display: boolean } | { readonly kind: 'edit'; readonly pos: number }

/** O que o editor devolve ao OK. O `omml` vai nulo: o sidecar o refaz do MathML. */
export interface EquationContent {
  readonly latex: string
  readonly mathml: string
  readonly display: boolean
}

/** A posição da equação selecionada (seleção de nó), ou `null`. */
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

/** Troca a equação em `pos`; devolve falso quando ali não há mais equação. */
export function replaceEquation(editor: Editor, pos: number, content: EquationContent): boolean {
  const node = editor.state.doc.nodeAt(pos)
  if (node === null || node.type.name !== 'math') return false

  const jc = content.display === node.attrs['display'] ? (node.attrs['jc'] as string | null) : null
  // As marcas de revisão são da equação antiga; o controle põe as da nova.
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
