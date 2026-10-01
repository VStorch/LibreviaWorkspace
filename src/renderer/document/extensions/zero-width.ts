import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'

/**
 * As âncoras sem largura: as pontas dos marcadores (bookmark.ts) e dos
 * comentários (comment.ts).
 *
 * Os quatro nós são de linha, atômicos e vazios — na tela não ocupam lugar, mas
 * para o ProseMirror são uma posição a mais, e para o navegador um
 * `contenteditable=false` vazio. Tudo o que trata deles como "não texto" mora
 * aqui, para que marcador e comentário se comportem igual.
 */
const ANCHORS = new Set(['bookmarkStart', 'bookmarkEnd', 'commentStart', 'commentEnd'])

/** Se o nó é ponta de marcador ou de comentário. */
export function isZeroWidthAnchor(node: ProseMirrorNode | null | undefined): boolean {
  return node !== null && node !== undefined && ANCHORS.has(node.type.name)
}

/** A posição depois das âncoras que abrem o bloco em `pos` — o começo do texto dele. */
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

/**
 * Meta da transação cuja seleção fica onde foi posta: a de quem escolhe o trecho
 * entre as âncoras de propósito (ir ao comentário), para que digitar por cima
 * troque o texto e não leve o comentário.
 */
export const KEEP_SELECTION = 'zeroWidthKeepSelection'

/** Ver o `appendTransaction` de `ZeroWidthAnchors`: a seleção leva as âncoras encostadas nas pontas do texto. */
export function extendOverAnchors(state: EditorState): Transaction | null {
  const { selection } = state
  if (!(selection instanceof TextSelection) || selection.empty) return null
  const { $from, $to } = selection

  let before = $from.parent.childBefore($from.parentOffset)
  let offset = $from.parentOffset
  while (offset > 0 && isZeroWidthAnchor(before.node)) {
    offset = before.offset
    before = $from.parent.childBefore(offset)
  }

  let after = $to.parent.childAfter($to.parentOffset)
  let end = $to.parentOffset
  while (end < $to.parent.content.size && isZeroWidthAnchor(after.node)) {
    end = after.offset + after.node!.nodeSize
    after = $to.parent.childAfter(end)
  }

  // Só o parágrafo cujo texto a seleção cobre inteiro leva as âncoras: o último
  // caractere selecionado (Shift+←) ou a última palavra não são o parágrafo, e
  // apagá-los não pode levar o comentário que termina ali.
  const sameBlock = $from.sameParent($to)
  const startCovered = offset === 0 && (!sameBlock || end === $to.parent.content.size)
  const endCovered = end === $to.parent.content.size && (!sameBlock || offset === 0)
  const from = startCovered ? $from.start() : selection.from
  const to = endCovered ? $to.end() : selection.to

  if (from === selection.from && to === selection.to) return null
  const [anchor, head] = selection.anchor <= selection.head ? [from, to] : [to, from]
  return state.tr.setSelection(TextSelection.create(state.doc, anchor, head))
}

/**
 * O cursor vazio depois (Backspace) ou antes (Delete) de âncoras, levado para o
 * outro lado delas. Devolve a posição nova, ou `null` quando não há âncora ali.
 */
export function pastAnchors(state: EditorState, direction: -1 | 1): number | null {
  const { selection } = state
  if (!(selection instanceof TextSelection) || !selection.empty) return null
  const $cursor = selection.$head
  let pos = $cursor.pos
  const limit = direction < 0 ? $cursor.start() : $cursor.end()
  while (pos !== limit) {
    const $pos = state.doc.resolve(pos)
    if (!isZeroWidthAnchor(direction < 0 ? $pos.nodeBefore : $pos.nodeAfter)) break
    pos += direction
  }
  return pos === $cursor.pos ? null : pos
}

/** A ponta móvel da seleção levada para depois das âncoras na direção dada, ou `null`. */
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

/**
 * O que o ProseMirror deixa para o `selectionchange` ler, lido já.
 *
 * O Chrome dispara o `selectionchange` numa tarefa à parte, e a tecla seguinte
 * pode chegar antes dele: `End` e `Shift+Home` depressa num parágrafo terminado
 * por âncora deixavam a linha selecionada na tela e o cursor parado no fim no
 * estado. O `Backspace` então agia sobre o estado velho — e, com um nó que não é
 * texto antes do cursor, o ProseMirror o apaga ele mesmo em vez de deixar o
 * navegador apagar a seleção: sumia a âncora (e o comentário com ela), e só no
 * segundo `Backspace` a linha. O `keydown` do ProseMirror esvazia só as mutações
 * pendentes, não a seleção; o `flush` faz o que o `selectionchange` faria. É o
 * mesmo atraso do comando que chega do menu pelo IPC (`flushSelection`).
 */
export function readPendingSelection(view: EditorView): void {
  const observer = (view as unknown as { domObserver?: { flush?: () => void } }).domObserver
  observer?.flush?.()
}

/**
 * O cursor e a seleção em volta das âncoras sem largura.
 *
 * - A tecla lê a seleção da tela antes de agir (ver `readPendingSelection`).
 * - `Backspace` e `Delete` sem seleção passam por cima das âncoras e apagam o
 *   caractere de verdade: a âncora não se vê, e apagá-la seria um toque perdido
 *   que ainda leva o comentário ou o marcador embora.
 * - A seleção que chega ao começo ou ao fim do texto de um parágrafo leva junto
 *   as âncoras encostadas ali.
 */
export const ZeroWidthAnchors = Extension.create({
  name: 'zeroWidthAnchors',
  // Antes do mapa de teclas do Tiptap, que trataria o `Backspace` com o estado velho.
  priority: 1000,

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('zeroWidthAnchors'),
        /**
         * Os nós não têm largura, e o navegador põe o cursor do lado de dentro
         * deles: `Shift+Home` numa legenda selecionava "Figura 1 — texto" sem o
         * começo do `_Ref` que a referência de página cita, e recortar e colar a
         * legenda deixava o marcador para trás — o F9 seguinte escrevia "Erro!
         * Indicador não definido.". Estendida, a seleção leva a âncora com o
         * texto, como no Word — e o comentário recortado com o parágrafo volta
         * com ele na colagem.
         */
        appendTransaction: (transactions, _old, state) =>
          transactions.some((tr) => tr.selectionSet) &&
          !transactions.some((tr) => tr.getMeta(KEEP_SELECTION) === true)
            ? extendOverAnchors(state)
            : null,

        props: {
          handleKeyDown(view, event) {
            if (event.isComposing || event.keyCode === 229) return false
            readPendingSelection(view)
            if (event.shiftKey && (event.key === 'ArrowLeft' || event.key === 'ArrowRight')) {
              // A ponta da seleção pula as âncoras antes de o navegador estendê-la:
              // Shift+← no fim de "Fim." depois de um comentário de ponto
              // selecionava só a âncora, e o Delete seguinte levava o comentário.
              const head = headPastAnchors(view.state, event.key === 'ArrowLeft' ? -1 : 1)
              if (head !== null) {
                // Sem seleção ainda, o cursor inteiro passa: a âncora não fica dentro dela.
                const anchor = view.state.selection.empty ? head : view.state.selection.anchor
                view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, anchor, head)))
              }
              return false
            }
            if (event.key === 'Enter') {
              // Enter com só âncoras antes do cursor: o parágrafo novo nasce
              // antes delas, e elas seguem com o texto em vez de ficarem sozinhas
              // na linha vazia de cima.
              const back = pastAnchors(view.state, -1)
              if (back !== null && back === view.state.selection.$head.start())
                view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, back)))
              return false
            }
            if (event.key !== 'Backspace' && event.key !== 'Delete') return false
            const pos = pastAnchors(view.state, event.key === 'Backspace' ? -1 : 1)
            // Só o cursor anda: o resto — o mapa de teclas, ou o próprio navegador
            // apagando o caractere — segue do lugar novo.
            if (pos !== null)
              view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)))
            return false
          },
        },
      }),
    ]
  },
})
