import type { Editor, NodeViewRendererProps } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { baseKeymap, toggleMark } from '@tiptap/pm/commands'
import { closeHistory } from '@tiptap/pm/history'
import { keymap } from '@tiptap/pm/keymap'
import {
  EditorState,
  Plugin,
  Selection,
  TextSelection,
  type Command,
  type Transaction,
} from '@tiptap/pm/state'
import { StepMap } from '@tiptap/pm/transform'
import { Decoration, DecorationSet, EditorView, type DecorationSource, type NodeView } from '@tiptap/pm/view'
import {
  SKIP_TRACKING,
  shouldTrack,
  trackTransaction,
  trackedDeleteKey,
  type TrackInputOptions,
} from './track-input.js'
import { revisionViewOf } from './revision-view.js'
import { drawsNoteNumber, noteRefAround } from './note-ref.js'

/**
 * O padrão de notas do ProseMirror: o corpo ganha um `EditorView` próprio, com o
 * nó como documento, e o que se digita vira passo do editor de fora, deslocado
 * para dentro (`pos + 1`) — um histórico só. O que muda por fora volta pela
 * diferença entre os corpos. O elemento do corpo mora fora do `contenteditable`
 * de fora: na área de notas da folha, ou num depósito escondido com a largura da
 * coluna, onde a paginação o mede; assim o IME trabalha num editor comum.
 */

const FROM_OUTSIDE = 'noteBody:fromOutside'

/** Clicar no número volta à referência. */
export const NOTE_NUMBER_CLASS = 'note-number'

let counter = 0
const bodies = new Map<string, NoteBody>()
const byReference = new WeakMap<Node, NoteBody>()
const pools = new WeakMap<EditorView, HTMLElement>()
const listeners = new Set<() => void>()
/** Até o texto pegar o foco de volta — ver `activeNoteOf`. */
const lastActive = new WeakMap<EditorView, NoteBody>()
const watched = new WeakSet<EditorView>()

function changed(): void {
  for (const listener of listeners) listener()
}

export function subscribeNoteBodies(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

export function noteBodyOf(reference: Node | null | undefined): NoteBody | undefined {
  return reference === null || reference === undefined ? undefined : byReference.get(reference)
}

export function noteBody(key: string): NoteBody | undefined {
  return bodies.get(key)
}

/** Para refazer o "editável" quando o de fora muda. */
export function noteBodiesOf(outer: EditorView): NoteBody[] {
  return [...bodies.values()].filter((body) => body.outer === outer)
}

/** Os corpos sem folha, inclusive no modo de leitura. Os que nasceram antes do depósito vão para lá agora. */
export function setNotePool(outer: EditorView, pool: HTMLElement | null): void {
  if (pool === null) {
    pools.delete(outer)
    return
  }
  pools.set(outer, pool)
  for (const body of noteBodiesOf(outer)) {
    if (body.body.parentElement !== null) continue
    pool.appendChild(body.body)
    if (body.pendingFocus) {
      body.pendingFocus = false
      body.focusEnd()
    }
  }
}

export function parkNoteBody(body: NoteBody, from: HTMLElement): void {
  if (body.body.parentElement !== from) return
  const pool = pools.get(body.outer)
  if (pool === undefined) return
  move(body, pool)
}

export function placeNoteBody(body: NoteBody, slot: HTMLElement): void {
  if (body.body.parentElement !== slot) move(body, slot)
  if (body.pendingFocus) {
    body.pendingFocus = false
    body.focusEnd()
  } else if (body.view.hasFocus()) {
    body.body.scrollIntoView({ block: 'nearest' })
  }
}

/** Mudar o elemento de lugar tira o foco: ele espera o corpo chegar à folha nova. */
function move(body: NoteBody, target: HTMLElement): void {
  const focused = body.view.hasFocus() || body.refocus
  target.appendChild(body.body)
  if (!focused) return
  if (target === pools.get(body.outer)) {
    body.refocus = true
    setTimeout(() => {
      body.refocus = false
    }, 0)
    return
  }
  body.refocus = false
  body.view.focus()
}

/** A que tem o foco, ou a última que o teve: o clique no menu de contexto tira o foco do corpo. */
export function activeNoteOf(outer: EditorView): NoteBody | null {
  const body = lastActive.get(outer)
  if (body === undefined || bodies.get(body.key) !== body || body.position() === undefined) return null
  return body
}

/** Em posições do documento de fora. */
export function caretOf(outer: EditorView): { from: number; to: number; note: NoteBody | null } {
  const note = activeNoteOf(outer)
  const at = note?.position()
  if (note !== null && at !== undefined) {
    const { from, to } = note.view.state.selection
    return { from: at + 1 + from, to: at + 1 + to, note }
  }
  const { from, to } = outer.state.selection
  return { from, to, note: null }
}

/**
 * O texto fica com o cursor depois da referência, e o corpo seleciona o trecho;
 * `false` fora de nota. `focus` leva o teclado ao corpo (o Próximo do menu).
 */
export function selectInNote(outer: EditorView, from: number, to: number, focus: boolean): boolean {
  const { doc } = outer.state
  const reference = noteRefAround(doc, from)
  const node = reference === null ? null : doc.nodeAt(reference)
  if (reference === null || node === null) return false
  const after = TextSelection.create(doc, reference + node.nodeSize)
  outer.dispatch(outer.state.tr.setSelection(after).scrollIntoView())
  const body = noteBodyOf(outer.nodeDOM(reference))
  if (body === undefined) return true
  if (focus) body.view.focus()
  body.select(from - reference - 1, to - reference - 1)
  return true
}

/** O corpo sem folha, no depósito, cede à referência. */
export function coordsInDocument(outer: EditorView, pos: number): { top: number; left: number } {
  const reference = noteRefAround(outer.state.doc, pos)
  if (reference === null) return outer.coordsAtPos(pos)
  const body = noteBodyOf(outer.nodeDOM(reference))
  if (body !== undefined && body.body.isConnected && body.body.parentElement !== pools.get(outer)) {
    const inner = Math.min(Math.max(pos - reference - 1, 0), body.view.state.doc.content.size)
    return body.view.coordsAtPos(inner)
  }
  return outer.coordsAtPos(reference)
}

export function flushNoteSelection(outer: EditorView): void {
  for (const body of noteBodiesOf(outer)) if (body.view.hasFocus()) syncSelection(body.view)
}

/** API interna do ProseMirror. */
function syncSelection(view: EditorView): void {
  const observer = (view as unknown as { domObserver?: { forceFlush?: () => void; flush?: () => void } })
    .domObserver
  observer?.forceFlush?.()
  observer?.flush?.()
}

function labelOf(decorations: readonly Decoration[]): string {
  for (const decoration of decorations) {
    const label: unknown = (decoration.spec as { noteLabel?: unknown }).noteLabel
    if (typeof label === 'string') return label
  }
  return ''
}

export class NoteBody implements NodeView {
  readonly key: string
  readonly dom: HTMLElement
  readonly body: HTMLElement
  readonly view: EditorView
  readonly outer: EditorView
  /** Foco pedido antes de o corpo ter folha: dado quando ele chega a uma. */
  pendingFocus = false
  refocus = false
  /** A nota recém-inserida e o que se digita nela se desfazem em dois tempos, como no Word. */
  separateHistory = false
  private node: ProseMirrorNode
  private label: string
  private innerDecorations: DecorationSource
  private numbers: { doc: ProseMirrorNode; label: string; set: DecorationSet } | null = null

  constructor(
    props: Pick<NodeViewRendererProps, 'node' | 'view' | 'getPos' | 'decorations' | 'innerDecorations'>,
    private readonly editor: Editor,
  ) {
    counter += 1
    this.key = `note-${counter}`
    this.node = props.node
    this.outer = props.view
    this.getPos = props.getPos
    this.label = labelOf(props.decorations)
    this.innerDecorations = props.innerDecorations

    this.dom = document.createElement('sup')
    this.render()
    this.dom.addEventListener('dblclick', (event) => {
      event.preventDefault()
      this.reveal()
    })

    const mount = document.createElement('div')
    this.view = new EditorView(
      { mount },
      {
        state: EditorState.create({ doc: props.node, plugins: this.plugins() }),
        dispatchTransaction: (tr) => this.dispatchInner(tr),
        editable: () => this.outer.editable,
        attributes: () => ({
          // É pela classe do modo que o CSS esconde o excluído ou o inserido.
          class: `page__content note-body revisions-${revisionViewOf(this.outer.state)}`,
          'data-note-kind': String(this.node.attrs['kind']),
          'data-note-key': this.key,
          spellcheck: this.outer.dom.getAttribute('spellcheck') ?? 'false',
        }),
        handleDOMEvents: {
          // O desfazer do menu do Electron chega como `beforeinput`, e não como tecla.
          beforeinput: (_view, event) => {
            const input = event as InputEvent
            if (input.inputType !== 'historyUndo' && input.inputType !== 'historyRedo') return false
            event.preventDefault()
            if (input.inputType === 'historyUndo') this.editor.commands.undo()
            else this.editor.commands.redo()
            return true
          },
          // O `selectionchange` das setas às vezes chega ~20 ms depois, e a tecla seguinte leria a seleção de antes.
          keydown: (view) => {
            syncSelection(view)
            return false
          },
          mousedown: (_view, event) => {
            if (!(event.target instanceof Element) || event.target.closest(`.${NOTE_NUMBER_CLASS}`) === null)
              return false
            event.preventDefault()
            this.revealReference()
            return true
          },
        },
      },
    )
    this.body = this.view.dom
    this.body.addEventListener('focus', () => lastActive.set(this.outer, this))
    if (!watched.has(this.outer)) {
      watched.add(this.outer)
      this.outer.dom.addEventListener('focus', () => lastActive.delete(this.outer))
    }
    bodies.set(this.key, this)
    byReference.set(this.dom, this)
    pools.get(this.outer)?.appendChild(this.body)
    changed()
  }

  private readonly getPos: () => number | undefined

  position(): number | undefined {
    return this.getPos()
  }

  /** Na ordem do `renderHTML` do nó, para o HTML não divergir. */
  private render(): void {
    const { kind, nid, mark } = this.node.attrs as { kind: unknown; nid: unknown; mark: unknown }
    this.dom.className = 'note-ref'
    this.dom.setAttribute('data-note-ref', '')
    this.dom.setAttribute('data-kind', String(kind))
    this.dom.setAttribute('data-note-key', this.key)
    if (typeof nid === 'string') this.dom.setAttribute('data-nid', nid)
    else this.dom.removeAttribute('data-nid')
    if (typeof mark === 'string' && mark !== '') this.dom.setAttribute('data-mark', mark)
    else this.dom.removeAttribute('data-mark')
    this.dom.textContent = typeof mark === 'string' ? mark : ''
  }

  private plugins(): Plugin[] {
    const schema = this.node.type.schema
    const redo: Command = () => this.editor.commands.redo()
    const marks: Record<string, Command> = {}
    for (const [key, name] of [
      ['Mod-b', 'bold'],
      ['Mod-i', 'italic'],
      ['Mod-u', 'underline'],
    ] as const) {
      const type = schema.marks[name]
      if (type !== undefined) marks[key] = toggleMark(type)
    }
    return [
      // Um histórico só, o do documento.
      keymap({ 'Mod-z': () => this.editor.commands.undo(), 'Mod-y': redo, 'Shift-Mod-z': redo, ...marks }),
      new Plugin({
        props: {
          handleKeyDown: (view, event) => trackedDeleteKey(view, event, () => this.tracking() !== null),
          decorations: () => this.innerDecorations,
        },
      }),
      new Plugin({ props: { decorations: (state) => this.numberDecoration(state.doc) } }),
      keymap(baseKeymap),
    ]
  }

  /** No começo do primeiro parágrafo, como o Word escreve. */
  private numberDecoration(doc: ProseMirrorNode): DecorationSet {
    if (this.numbers?.doc === doc && this.numbers.label === this.label) return this.numbers.set
    const label = this.label
    const at = doc.firstChild?.isTextblock === true ? 1 : 0
    const set =
      label === '' || !drawsNoteNumber(this.node)
        ? DecorationSet.empty
        : DecorationSet.create(doc, [
            Decoration.widget(
              at,
              () => {
                const number = document.createElement('span')
                number.className = NOTE_NUMBER_CLASS
                number.contentEditable = 'false'
                number.textContent = label
                return number
              },
              { side: -1, ignoreSelection: true, key: `note-number:${label}` },
            ),
          ])
    this.numbers = { doc, label, set }
    return set
  }

  private tracking(): TrackInputOptions | null {
    const extension = this.editor.extensionManager.extensions.find((item) => item.name === 'trackInput')
    const options = extension?.options as TrackInputOptions | undefined
    return options?.isTracking() === true ? options : null
  }

  private dispatchInner(tr: Transaction): void {
    if (tr.getMeta(FROM_OUTSIDE) === true || !tr.docChanged) {
      this.view.updateState(this.view.state.apply(tr))
      if (tr.docChanged) changed()
      return
    }

    const at = this.syncedPosition()
    if (at === null) return

    // Aqui, e não no `dispatchTransaction` de fora: só aqui existe a seleção do corpo.
    const options = this.tracking()
    const { state, transactions } = this.view.state.applyTransaction(this.trackedOrAsIs(tr, options))
    this.view.updateState(state)
    changed()
    this.forwardOutside(transactions, at, { tracked: options !== null, source: tr })
  }

  /**
   * Se um passo anterior não chegou lá fora, este cairia deslocado: o corpo
   * volta ao que o documento tem, e perder uma tecla é melhor que escrever no
   * lugar errado.
   */
  private syncedPosition(): number | null {
    const at = this.getPos()
    const current = at === undefined ? null : this.outer.state.doc.nodeAt(at)
    if (at !== undefined && current !== null && current.content.eq(this.view.state.doc.content)) return at
    if (current !== null) {
      const { doc } = this.view.state
      this.view.dispatch(
        this.view.state.tr.replace(0, doc.content.size, current.slice(0)).setMeta(FROM_OUTSIDE, true),
      )
    }
    return null
  }

  private trackedOrAsIs(tr: Transaction, options: TrackInputOptions | null): Transaction {
    if (options === null || !shouldTrack(tr)) return tr
    try {
      return trackTransaction(tr, this.view.state, options.author(), new Date(), {
        composing: this.view.composing,
      })
    } catch (error) {
      // Melhor a edição sem controle que a edição perdida.
      console.error(error)
      return tr
    }
  }

  private forwardOutside(
    transactions: readonly Transaction[],
    at: number,
    { tracked, source }: { readonly tracked: boolean; readonly source: Transaction },
  ): void {
    const outer = this.outer.state.tr
    const offset = StepMap.offset(at + 1)
    for (const step of transactions.flatMap((transaction) => transaction.steps)) {
      const mapped = step.map(offset)
      if (mapped !== null) outer.step(mapped)
    }
    if (!outer.docChanged) return
    if (tracked) outer.setMeta(SKIP_TRACKING, true)
    if (this.separateHistory) {
      this.separateHistory = false
      closeHistory(outer)
    }
    if (source.getMeta('addToHistory') === false) outer.setMeta('addToHistory', false)
    this.outer.dispatch(outer)
  }

  update(
    node: ProseMirrorNode,
    decorations: readonly Decoration[],
    innerDecorations: DecorationSource,
  ): boolean {
    if (node.type !== this.node.type) return false
    const markup = !node.sameMarkup(this.node)
    const label = labelOf(decorations)
    const redraw = label !== this.label || innerDecorations !== this.innerDecorations
    this.node = node
    this.label = label
    this.innerDecorations = innerDecorations
    if (markup) this.render()

    const state = this.view.state
    const start = node.content.findDiffStart(state.doc.content)
    if (start !== null) {
      const end = node.content.findDiffEnd(state.doc.content)
      let endA = end?.a ?? node.content.size
      let endB = end?.b ?? state.doc.content.size
      const overlap = start - Math.min(endA, endB)
      if (overlap > 0) {
        endA += overlap
        endB += overlap
      }
      this.view.dispatch(state.tr.replace(start, endB, node.slice(start, endA)).setMeta(FROM_OUTSIDE, true))
    } else if (redraw || markup) {
      this.view.setProps({})
    }
    return true
  }

  refresh(): void {
    this.view.setProps({})
  }

  /** O corpo ainda no depósito já recebe o foco, para o que se digita não ir parar no texto. */
  reveal(): void {
    if (!this.body.isConnected) {
      this.pendingFocus = true
      return
    }
    this.focusEnd()
  }

  focusEnd(): void {
    const { state } = this.view
    this.view.dispatch(state.tr.setSelection(Selection.atEnd(state.doc)).setMeta(FROM_OUTSIDE, true))
    this.view.focus()
    if (this.body.parentElement !== pools.get(this.outer)) this.body.scrollIntoView({ block: 'nearest' })
  }

  select(from: number, to: number): void {
    const { state } = this.view
    const size = state.doc.content.size
    const selection = TextSelection.create(state.doc, Math.min(from, size), Math.min(to, size))
    this.view.dispatch(state.tr.setSelection(selection).setMeta(FROM_OUTSIDE, true))
    this.body.scrollIntoView({ block: 'nearest' })
  }

  revealReference(): void {
    const pos = this.getPos()
    if (pos === undefined) return
    const { state } = this.outer
    const after = Math.min(pos + this.node.nodeSize, state.doc.content.size)
    this.outer.dispatch(state.tr.setSelection(TextSelection.create(state.doc, after)).scrollIntoView())
    this.outer.focus()
  }

  stopEvent(): boolean {
    return false
  }

  ignoreMutation(): boolean {
    return true
  }

  destroy(): void {
    // A nota com o cursor saiu do documento: o foco volta ao texto, numa
    // microtask, porque o editor de fora ainda está no meio da atualização.
    const focused = this.view.hasFocus() || this.refocus
    if (focused)
      queueMicrotask(() => {
        if (!this.outer.isDestroyed) this.outer.focus()
      })
    this.view.destroy()
    this.body.remove()
    bodies.delete(this.key)
    changed()
  }
}

export function noteRefView(editor: Editor) {
  return (props: NodeViewRendererProps): NoteBody => new NoteBody(props, editor)
}
