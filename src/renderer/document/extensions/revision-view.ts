import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode, ResolvedPos } from '@tiptap/pm/model'
import {
  Plugin,
  PluginKey,
  Selection,
  TextSelection,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { RevisionView } from '@shared/types.js'
import { DELETION, INSERTION, ZERO_WIDTH, blockRevisionOf } from './track-changes.js'

/**
 * Controle de alterações (M10, fase 3): como a janela mostra as alterações.
 *
 * - **Marcação completa**: tudo à vista, como a fase 1 desenha.
 * - **Marcação simples**: o texto final — o excluído some, o inserido fica sem
 *   marca — e uma barra na margem ao lado do parágrafo alterado.
 * - **Sem marcação**: o texto final, limpo.
 * - **Original**: o texto de antes — o inserido some, o excluído fica sem marca.
 *   O editor fica travado nesse modo (ver DocumentEditor).
 *
 * Quem esconde é o CSS, por uma classe na raiz do editor (`revisions-<modo>`); as
 * decorações marcam os blocos que somem inteiros e os parágrafos alterados. O
 * documento não muda: trocar de modo não suja o arquivo nem entra no desfazer.
 *
 * Esconder não basta: o cursor não pode cair dentro do texto escondido, onde se
 * digitaria sem ver. A seleção que para lá é empurrada para a borda, no sentido
 * em que andava, e o Backspace e o Delete passam por cima do escondido antes de
 * apagar o caractere que se vê.
 */

interface ViewState {
  readonly view: RevisionView
  readonly decorations: DecorationSet
}

export const revisionViewKey = new PluginKey<ViewState>('revisionView')

/** O modo do editor. */
export function revisionViewOf(state: EditorState): RevisionView {
  return revisionViewKey.getState(state)?.view ?? RevisionView.All
}

/** A transação que troca o modo — sem mudar o documento. */
export function setRevisionViewMeta(tr: Transaction, view: RevisionView): Transaction {
  return tr.setMeta(revisionViewKey, view)
}

/** O nó em linha some neste modo? */
export function isHiddenInline(node: ProseMirrorNode, view: RevisionView): boolean {
  if (view === RevisionView.All) return false
  const name = view === RevisionView.Original ? INSERTION : DELETION
  return node.marks.some((mark) => mark.type.name === name)
}

/** A revisão de bloco (marca de parágrafo, linha) some neste modo? */
function isHiddenRevision(value: unknown, view: RevisionView): boolean {
  if (view === RevisionView.All) return false
  const kind = blockRevisionOf(value)?.kind
  return view === RevisionView.Original ? kind === 'ins' : kind === 'del'
}

/** Todo o conteúdo do bloco de texto some — e há o que sumir. */
function allContentHidden(block: ProseMirrorNode, view: RevisionView): boolean {
  let hidden = false
  let visible = false
  block.forEach((child) => {
    if (isHiddenInline(child, view)) hidden = true
    else if (!ZERO_WIDTH.has(child.type.name)) visible = true
  })
  return hidden && !visible
}

/**
 * O bloco some inteiro neste modo: a linha de tabela da revisão escondida, e o
 * parágrafo cuja marca some junto com todo o texto (o parágrafo inserido inteiro,
 * no Original; o excluído inteiro, no texto final).
 */
export function isHiddenBlock(node: ProseMirrorNode, view: RevisionView): boolean {
  if (view === RevisionView.All) return false
  if (node.type.name === 'tableRow') return isHiddenRevision(node.attrs['rowRevision'], view)
  if (!node.isTextblock || !isHiddenRevision(node.attrs['markRevision'], view)) return false
  return node.content.size === 0 || allContentHidden(node, view)
}

/**
 * Os trechos escondidos do bloco de texto, em posições relativas a ele: de onde
 * começa o primeiro nó escondido até onde termina o último, atravessando as
 * pontas sem largura entre eles.
 */
export function hiddenRuns(block: ProseMirrorNode, view: RevisionView): { from: number; to: number }[] {
  const runs: { from: number; to: number }[] = []
  if (view === RevisionView.All || !block.isTextblock) return runs
  let run: { from: number; to: number } | null = null
  block.forEach((child, offset) => {
    if (isHiddenInline(child, view)) {
      if (run === null) {
        run = { from: offset, to: offset + child.nodeSize }
        runs.push(run)
      } else {
        run.to = offset + child.nodeSize
      }
    } else if (!ZERO_WIDTH.has(child.type.name)) {
      run = null
    }
  })
  return runs
}

/** Uma posição fora do escondido, andando no sentido `dir` (ou, sem saída, no outro). */
export function visiblePosition(doc: ProseMirrorNode, pos: number, view: RevisionView, dir: 1 | -1): number {
  if (view === RevisionView.All) return pos
  // Um bloco escondido pode levar a outro, e a saída para trás de volta ao primeiro.
  // A gangorra se reconhece pela posição repetida — não por um número de voltas,
  // que cortava no meio uma sequência longa de parágrafos excluídos.
  const seen = new Set<number>()
  while (!seen.has(pos)) {
    seen.add(pos)
    const $pos = doc.resolve(pos)
    const outside = outOfHiddenBlock(doc, $pos, view, dir)
    if (outside === null) return pos
    if (outside !== pos) {
      pos = outside
      continue
    }
    const start = $pos.start()
    for (const run of hiddenRuns($pos.parent, view)) {
      if (run.from < $pos.parentOffset && $pos.parentOffset < run.to) {
        return start + (dir > 0 ? run.to : run.from)
      }
    }
    return pos
  }
  return pos
}

/** A posição seguinte fora do bloco escondido em volta de `$pos`, a mesma se não há, ou `null` sem saída. */
function outOfHiddenBlock(
  doc: ProseMirrorNode,
  $pos: ResolvedPos,
  view: RevisionView,
  dir: 1 | -1,
): number | null {
  for (let depth = 1; depth <= $pos.depth; depth++) {
    if (!isHiddenBlock($pos.node(depth), view)) continue
    const edge = dir > 0 ? $pos.after(depth) : $pos.before(depth)
    const found =
      Selection.findFrom(doc.resolve(edge), dir, true) ??
      Selection.findFrom(doc.resolve(dir > 0 ? $pos.before(depth) : $pos.after(depth)), -dir as 1 | -1, true)
    return found === null ? null : found.head
  }
  return $pos.pos
}

/**
 * Apaga de `from` a `to` só o que o modo deixa ver: o texto escondido no meio
 * fica (o excluído que não se vê não é apagado de verdade por uma seleção que
 * passou por cima dele sem mostrá-lo).
 */
export function deleteVisible(tr: Transaction, from: number, to: number, view: RevisionView): Transaction {
  const hidden: { from: number; to: number }[] = []
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isInline && isHiddenInline(node, view)) {
      hidden.push({ from: Math.max(pos, from), to: Math.min(pos + node.nodeSize, to) })
    }
    return true
  })
  const pieces: { from: number; to: number }[] = []
  let cursor = from
  for (const range of hidden) {
    if (range.from > cursor) pieces.push({ from: cursor, to: range.from })
    cursor = Math.max(cursor, range.to)
  }
  if (cursor < to) pieces.push({ from: cursor, to })
  // De trás para frente: apagar adiante não mexe nas posições de trás.
  for (const piece of pieces.reverse()) tr.delete(piece.from, piece.to)
  return tr
}

/** A seleção de texto empurrada para fora do escondido — ou `null`, se já está fora. */
export function visibleSelection(
  doc: ProseMirrorNode,
  selection: Selection,
  view: RevisionView,
  dir: 1 | -1,
): Selection | null {
  if (view === RevisionView.All || !(selection instanceof TextSelection)) return null
  const head = visiblePosition(doc, selection.head, view, dir)
  const anchor = selection.empty ? head : visiblePosition(doc, selection.anchor, view, dir)
  if (head === selection.head && anchor === selection.anchor) return null
  return TextSelection.between(doc.resolve(anchor), doc.resolve(head))
}

function buildDecorations(doc: ProseMirrorNode, view: RevisionView): DecorationSet {
  if (view === RevisionView.All) return DecorationSet.empty
  const decorations: Decoration[] = []
  doc.descendants((node, pos) => {
    if (isHiddenBlock(node, view)) {
      decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: 'revision-hidden' }))
      return false
    }
    if (!node.isTextblock) return true
    const classes: string[] = []
    // O texto todo escondido, a marca de parágrafo não: fica a linha vazia.
    if (node.content.size > 0 && allContentHidden(node, view)) classes.push('revision-blank')
    if (view === RevisionView.Simple && hasRevision(node)) classes.push('revision-changed')
    if (classes.length > 0) {
      decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: classes.join(' ') }))
    }
    return false
  })
  return DecorationSet.create(doc, decorations)
}

function hasRevision(block: ProseMirrorNode): boolean {
  if (blockRevisionOf(block.attrs['markRevision']) !== null) return true
  let found = false
  block.forEach((child) => {
    if (child.marks.some((mark) => mark.type.name === INSERTION || mark.type.name === DELETION)) found = true
  })
  return found
}

export const RevisionViewExtension = Extension.create({
  name: 'revisionView',
  // Antes do controle do que se digita (track-input.ts), que apagaria o
  // escondido vizinho; depois das âncoras sem largura (zero-width.ts, 1000).
  priority: 900,

  addProseMirrorPlugins() {
    return [
      new Plugin<ViewState>({
        key: revisionViewKey,
        state: {
          init: () => ({ view: RevisionView.All, decorations: DecorationSet.empty }),
          apply(tr, value, _old, state) {
            const chosen = tr.getMeta(revisionViewKey) as RevisionView | undefined
            if (chosen !== undefined && chosen !== value.view) {
              return { view: chosen, decorations: buildDecorations(state.doc, chosen) }
            }
            if (!tr.docChanged || value.view === RevisionView.All) return value
            return { view: value.view, decorations: buildDecorations(state.doc, value.view) }
          },
        },
        props: {
          attributes: (state) => ({ class: `revisions-${revisionViewOf(state)}` }),
          decorations: (state) => revisionViewKey.getState(state)?.decorations ?? null,
          // Com algo escondido, o Backspace e o Delete são feitos aqui, e não pelo
          // navegador: apagando sozinho, ele levava junto o excluído escondido
          // (`display:none`) ou o relia como tachado comum. Apaga-se só o que se
          // vê — um caractere, ou a parte visível da seleção —, e o escondido fica.
          handleKeyDown(view, event) {
            if (event.key !== 'Backspace' && event.key !== 'Delete') return false
            if (event.ctrlKey || event.metaKey || event.altKey || event.shiftKey || view.composing)
              return false
            const mode = revisionViewOf(view.state)
            const { selection } = view.state
            if (mode === RevisionView.All || !(selection instanceof TextSelection)) return false
            if (!selection.empty) {
              view.dispatch(deleteVisible(view.state.tr, selection.from, selection.to, mode).scrollIntoView())
              return true
            }
            const backward = event.key === 'Backspace'
            const $cursor = selection.$head
            const start = $cursor.start()
            let offset = $cursor.parentOffset
            // Passa o escondido e as âncoras sem largura até o caractere que se vê.
            for (let moved = true; moved;) {
              moved = false
              for (const run of hiddenRuns($cursor.parent, mode)) {
                if (backward ? run.to === offset : run.from === offset) {
                  offset = backward ? run.from : run.to
                  moved = true
                }
              }
              const index = backward ? offset - 1 : offset
              const child =
                index >= 0 && index < $cursor.parent.content.size ? $cursor.parent.childAfter(index) : null
              if (child?.node !== null && child?.node !== undefined && ZERO_WIDTH.has(child.node.type.name)) {
                offset += backward ? -1 : 1
                moved = true
              }
            }
            const $at = view.state.doc.resolve(start + offset)
            const node = backward ? $at.nodeBefore : $at.nodeAfter
            if (node === null || !node.isText || node.text === undefined) {
              // Na borda do bloco: o cursor fica do lado de fora do escondido, e o
              // juntar de parágrafos segue o caminho de sempre.
              if (start + offset !== $cursor.pos)
                view.dispatch(
                  view.state.tr.setSelection(TextSelection.create(view.state.doc, start + offset)),
                )
              return false
            }
            const text = node.text
            const unit = backward ? text.charCodeAt(text.length - 1) : text.charCodeAt(0)
            const surrogate = backward ? unit >= 0xdc00 && unit <= 0xdfff : unit >= 0xd800 && unit <= 0xdbff
            const size = surrogate && text.length > 1 ? 2 : 1
            const from = backward ? start + offset - size : start + offset
            view.dispatch(view.state.tr.delete(from, from + size).scrollIntoView())
            return true
          },
          // Digitar por cima de uma seleção que passa por texto escondido: o
          // navegador apagaria a seleção inteira (e refaria o excluído como
          // tachado) antes do `handleTextInput` — então a tecla para já no
          // `beforeinput`, e só o que se vê é trocado.
          handleDOMEvents: {
            beforeinput(view, event) {
              const input = event as InputEvent
              const mode = revisionViewOf(view.state)
              const { selection } = view.state
              if (mode === RevisionView.All || selection.empty || view.composing) return false
              if (input.inputType !== 'insertText' || input.data === null) return false
              event.preventDefault()
              const tr = deleteVisible(view.state.tr, selection.from, selection.to, mode)
              tr.insertText(input.data, tr.mapping.map(selection.from, -1))
              view.dispatch(tr.scrollIntoView())
              return true
            },
          },
          handleTextInput(view, from, to, text) {
            const mode = revisionViewOf(view.state)
            if (mode === RevisionView.All || from === to || view.composing) return false
            const tr = deleteVisible(view.state.tr, from, to, mode)
            const at = tr.mapping.map(from, -1)
            tr.insertText(text, at)
            view.dispatch(tr.scrollIntoView())
            return true
          },
        },
        // A seleção que caiu no escondido — pela seta, pelo clique, pela edição
        // — sai para a borda, no sentido em que andava.
        appendTransaction(transactions, oldState, newState) {
          const mode = revisionViewOf(newState)
          if (mode === RevisionView.All) return null
          if (
            !transactions.some(
              (tr) => tr.selectionSet || tr.docChanged || tr.getMeta(revisionViewKey) !== undefined,
            )
          )
            return null
          const dir = newState.selection.head >= oldState.selection.head ? 1 : -1
          const selection = visibleSelection(newState.doc, newState.selection, mode, dir)
          return selection === null ? null : newState.tr.setSelection(selection)
        },
      }),
    ]
  },
})
