import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode, ResolvedPos } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'

/**
 * Pontas de marcador e de comentário: atômicas, vazias e sem largura na tela,
 * mas uma posição para o ProseMirror. O que as trata como "não texto" mora aqui.
 */
const ANCHORS = new Set(['bookmarkStart', 'bookmarkEnd', 'commentStart', 'commentEnd'])

export function isZeroWidthAnchor(node: ProseMirrorNode | null | undefined): boolean {
  return node !== null && node !== undefined && ANCHORS.has(node.type.name)
}

/** O começo do texto do bloco, depois das âncoras que o abrem. */
export function textStartOf(doc: ProseMirrorNode, pos: number): number {
  const block = doc.nodeAt(pos)
  let start = pos + 1
  for (let index = 0; block !== null && index < block.childCount; index++) {
    const child = block.child(index)
    if (!isZeroWidthAnchor(child)) break
    start += child.nodeSize
  }
  return start
}

/** O `keyCode` das teclas que o método de entrada ainda está compondo. */
const IME_PROCESS_KEY_CODE = 229

/** A seleção posta de propósito entre as âncoras (ir ao comentário): digitar troca o texto e não leva o comentário. */
export const KEEP_SELECTION = 'zeroWidthKeepSelection'

/** A seleção leva as âncoras encostadas nas pontas do texto. */
export function extendOverAnchors(state: EditorState): Transaction | null {
  const { selection } = state
  if (!(selection instanceof TextSelection) || selection.empty) return null
  const { $from, $to } = selection

  const offset = offsetBeforeAnchors($from)
  const end = offsetAfterAnchors($to)

  // Só o parágrafo que a seleção cobre inteiro leva as âncoras.
  const sameBlock = $from.sameParent($to)
  const startCovered = offset === 0 && (!sameBlock || end === $to.parent.content.size)
  const endCovered = end === $to.parent.content.size && (!sameBlock || offset === 0)
  const from = startCovered ? $from.start() : selection.from
  const to = endCovered ? $to.end() : selection.to

  if (from === selection.from && to === selection.to) return null
  const [anchor, head] = selection.anchor <= selection.head ? [from, to] : [to, from]
  return state.tr.setSelection(TextSelection.create(state.doc, anchor, head))
}

function offsetBeforeAnchors($from: ResolvedPos): number {
  let offset = $from.parentOffset
  let before = $from.parent.childBefore(offset)
  while (offset > 0 && isZeroWidthAnchor(before.node)) {
    offset = before.offset
    before = $from.parent.childBefore(offset)
  }
  return offset
}

function offsetAfterAnchors($to: ResolvedPos): number {
  let end = $to.parentOffset
  let after = $to.parent.childAfter(end)
  while (end < $to.parent.content.size && isZeroWidthAnchor(after.node)) {
    end = after.offset + after.node!.nodeSize
    after = $to.parent.childAfter(end)
  }
  return end
}

/** Backspace e Delete passam por cima das âncoras; `null` sem âncora ali. */
export function pastAnchors(state: EditorState, direction: -1 | 1): number | null {
  const { selection } = state
  if (!(selection instanceof TextSelection) || !selection.empty) return null
  return headPastAnchors(state, direction)
}

function headPastAnchors(state: EditorState, direction: -1 | 1): number | null {
  const { selection } = state
  if (!(selection instanceof TextSelection)) return null
  const $head = selection.$head
  let pos = $head.pos
  const limit = direction < 0 ? $head.start() : $head.end()
  while (pos !== limit) {
    const $pos = state.doc.resolve(pos)
    if (!isZeroWidthAnchor(direction < 0 ? $pos.nodeBefore : $pos.nodeAfter)) break
    pos += direction
  }
  return pos === $head.pos ? null : pos
}

/** O cursor ou a ponta da seleção pulam as âncoras antes de o navegador agir; a tecla segue adiante. */
function moveOverAnchors(view: EditorView, event: KeyboardEvent): void {
  if (event.shiftKey && (event.key === 'ArrowLeft' || event.key === 'ArrowRight')) {
    // A ponta pula as âncoras antes de o navegador estendê-la.
    const head = headPastAnchors(view.state, event.key === 'ArrowLeft' ? -1 : 1)
    if (head === null) return
    const anchor = view.state.selection.empty ? head : view.state.selection.anchor
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, anchor, head)))
    return
  }
  if (event.key === 'Enter') {
    // Enter com só âncoras antes do cursor: elas seguem com o texto.
    const back = pastAnchors(view.state, -1)
    if (back !== null && back === view.state.selection.$head.start()) moveCursor(view, back)
    return
  }
  if (event.key !== 'Backspace' && event.key !== 'Delete') return
  // Só o cursor anda; o resto segue do lugar novo.
  const pos = pastAnchors(view.state, event.key === 'Backspace' ? -1 : 1)
  if (pos !== null) moveCursor(view, pos)
}

function moveCursor(view: EditorView, pos: number): void {
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)))
}

/**
 * O Chrome dispara o `selectionchange` numa tarefa à parte, e a tecla seguinte
 * pode chegar antes: o Backspace agiria sobre o estado velho e apagaria a âncora,
 * e o comentário com ela. O `flush` faz o que o `selectionchange` faria.
 */
export function readPendingSelection(view: EditorView): void {
  const observer = (view as unknown as { domObserver?: { flush?: () => void } }).domObserver
  observer?.flush?.()
}

/**
 * Backspace e Delete passam por cima das âncoras, que não se veem: apagá-las
 * seria um toque perdido que ainda leva o comentário. A seleção que chega à ponta
 * do texto de um parágrafo leva as âncoras encostadas ali.
 */
export const ZeroWidthAnchors = Extension.create({
  name: 'zeroWidthAnchors',
  // Antes do mapa de teclas do Tiptap, que trataria o Backspace com o estado velho.
  priority: 1000,

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('zeroWidthAnchors'),
        /**
         * O navegador põe o cursor dentro das âncoras: `Shift+Home` numa legenda
         * deixaria o `_Ref` para trás. Estendida, a seleção leva a âncora com o
         * texto, como no Word.
         */
        appendTransaction: (transactions, _old, state) =>
          transactions.some((tr) => tr.selectionSet) &&
          !transactions.some((tr) => tr.getMeta(KEEP_SELECTION) === true)
            ? extendOverAnchors(state)
            : null,

        props: {
          handleKeyDown(view, event) {
            if (event.isComposing || event.keyCode === IME_PROCESS_KEY_CODE) return false
            readPendingSelection(view)
            moveOverAnchors(view, event)
            return false
          },
        },
      }),
    ]
  },
})
