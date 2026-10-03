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
 * O corpo da nota, editável no pé da página.
 *
 * O padrão do exemplo de notas do ProseMirror: a referência (`noteRef`) é
 * desenhada por esta vista, e o corpo dela ganha um `EditorView` próprio, com o
 * próprio nó como documento. O que se digita lá dentro vira passo do editor de
 * fora, deslocado para dentro do nó (`pos + 1`): há **um** histórico só, e
 * desfazer dentro da nota desfaz no documento. O que muda por fora — desfazer,
 * substituir tudo, aceitar uma alteração — volta para dentro pela diferença
 * entre os dois corpos.
 *
 * O elemento do corpo não fica dentro do `contenteditable` de fora: ele mora na
 * área de notas da folha (`PaperSheet`), e antes de a paginação o pôr lá, num
 * depósito escondido com a largura da coluna de texto — onde ele já tem altura,
 * que é o que a paginação mede. Com isso os eventos do corpo nunca passam pelo
 * editor de fora, e o IME trabalha num editor comum.
 */

/** A transação que só traz para dentro o que mudou fora. */
const FROM_OUTSIDE = 'noteBody:fromOutside'

/** O número desenhado no começo do corpo; clicar nele volta à referência. */
export const NOTE_NUMBER_CLASS = 'note-number'

let counter = 0
const bodies = new Map<string, NoteBody>()
const byReference = new WeakMap<Node, NoteBody>()
const pools = new WeakMap<EditorView, HTMLElement>()
const listeners = new Set<() => void>()
/** A última nota que teve o foco, até o texto pegá-lo de volta — ver `activeNoteOf`. */
const lastActive = new WeakMap<EditorView, NoteBody>()
const watched = new WeakSet<EditorView>()

function changed(): void {
  for (const listener of listeners) listener()
}

/** Para a cópia da continuação, que precisa se refazer quando o corpo muda. */
export function subscribeNoteBodies(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

/** O corpo cuja referência é este elemento (`view.nodeDOM` da referência). */
export function noteBodyOf(reference: Node | null | undefined): NoteBody | undefined {
  return reference === null || reference === undefined ? undefined : byReference.get(reference)
}

export function noteBody(key: string): NoteBody | undefined {
  return bodies.get(key)
}

/** Os corpos de um editor — para refazer o "editável" quando o de fora muda. */
export function noteBodiesOf(outer: EditorView): NoteBody[] {
  return [...bodies.values()].filter((body) => body.outer === outer)
}

/**
 * O depósito dos corpos que ainda não têm folha (ou que estão no modo de
 * leitura, sem folha nenhuma). Os que nasceram antes dele vão para lá agora.
 */
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

/** Devolve o corpo ao depósito — a folha que o mostrava saiu de cena. */
export function parkNoteBody(body: NoteBody, from: HTMLElement): void {
  if (body.body.parentElement !== from) return
  const pool = pools.get(body.outer)
  if (pool === undefined) return
  move(body, pool)
}

/** Põe o corpo na área de notas da folha. */
export function placeNoteBody(body: NoteBody, slot: HTMLElement): void {
  if (body.body.parentElement !== slot) move(body, slot)
  if (body.pendingFocus) {
    body.pendingFocus = false
    body.focusEnd()
  } else if (body.view.hasFocus()) {
    body.body.scrollIntoView({ block: 'nearest' })
  }
}

/**
 * Mudar o elemento de lugar tira o foco dele; quem estava digitando continua.
 * A folha que se refaz (o número dela mudou) devolve o corpo ao depósito e o
 * pega de volta no mesmo desenho: o foco espera o corpo chegar à folha nova.
 */
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

/**
 * A nota onde está o cursor: a que tem o foco, ou a última que o teve enquanto o
 * texto não o pegou de volta. O clique no menu de contexto tira o foco do corpo,
 * e o comando que ele dispara — aceitar a alteração, comentar — ainda é sobre a
 * nota em que a pessoa estava.
 */
export function activeNoteOf(outer: EditorView): NoteBody | null {
  const body = lastActive.get(outer)
  if (body === undefined || bodies.get(body.key) !== body || body.position() === undefined) return null
  return body
}

/**
 * A seleção de quem tem o cursor, em posições do documento de fora: a da nota
 * ativa, deslocada para dentro do nó (o corpo é o próprio nó como documento), ou
 * a do texto.
 */
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
 * Escolhe um trecho que mora dentro de uma nota: o texto fica com o cursor
 * depois da referência, e o corpo seleciona o trecho. Devolve `false` quando o
 * trecho não está numa nota — aí quem chamou o escolhe no texto. `focus` leva o
 * teclado para o corpo (o Próximo do menu); o clique no cartão do painel não.
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

/**
 * Onde uma posição do documento cai na tela. Dentro de uma nota, no corpo dela,
 * no pé da página; o corpo que ainda não tem folha (no depósito, escondido) cede
 * à referência, no texto.
 */
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

/** Lê já a seleção da nota que tem o foco — o comando do menu age sobre ela. */
export function flushNoteSelection(outer: EditorView): void {
  for (const body of noteBodiesOf(outer)) if (body.view.hasFocus()) syncSelection(body.view)
}

/** Lê já a seleção do DOM (API interna do ProseMirror; sem ela, nada muda). */
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
  /** A referência, no texto. */
  readonly dom: HTMLElement
  /** O corpo, editável, na área de notas. */
  readonly body: HTMLElement
  readonly view: EditorView
  readonly outer: EditorView
  /** Foco pedido antes de o corpo ter folha: dado quando ele chega a uma. */
  pendingFocus = false
  /** Estava com o foco quando foi para o depósito (ver `move`). */
  refocus = false
  /**
   * A próxima edição abre um passo novo de desfazer: a nota recém-inserida e o
   * que se digita nela se desfazem em dois tempos, como no Word.
   */
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
          // A classe do modo de mostrar as alterações, como no editor de fora:
          // é por ela que o CSS esconde o excluído ou o inserido.
          class: `page__content note-body revisions-${revisionViewOf(this.outer.state)}`,
          'data-note-kind': String(this.node.attrs['kind']),
          'data-note-key': this.key,
          spellcheck: this.outer.dom.getAttribute('spellcheck') ?? 'false',
        }),
        handleDOMEvents: {
          // O desfazer do menu (papel `undo` do Electron) chega como `beforeinput`,
          // e não como tecla: vai para o histórico do documento, como o atalho.
          beforeinput: (_view, event) => {
            const input = event as InputEvent
            if (input.inputType !== 'historyUndo' && input.inputType !== 'historyRedo') return false
            event.preventDefault()
            if (input.inputType === 'historyUndo') this.editor.commands.undo()
            else this.editor.commands.redo()
            return true
          },
          // A seleção que as setas mudaram chega ao estado por um `selectionchange`
          // que o observador do ProseMirror às vezes adia (~20 ms); uma tecla logo
          // depois — o Backspace controlado, o Enter — leria a seleção de antes.
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
    // A nota ativa (ver `activeNoteOf`): a que recebeu o foco por último, até o
    // texto recebê-lo.
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

  /** A posição da referência no documento de fora. */
  position(): number | undefined {
    return this.getPos()
  }

  /** A ordem dos atributos é a do `renderHTML` do nó, para o HTML não divergir. */
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
      // Um histórico só: o do documento. Desfazer aqui desfaz lá, e o que voltar
      // para o corpo chega por `update`.
      keymap({ 'Mod-z': () => this.editor.commands.undo(), 'Mod-y': redo, 'Shift-Mod-z': redo, ...marks }),
      new Plugin({
        props: {
          handleKeyDown: (view, event) => trackedDeleteKey(view, event, () => this.tracking() !== null),
          // As decorações de fora que caem dentro da nota: a busca, o modo de
          // mostrar as alterações, os caracteres sem largura.
          decorations: () => this.innerDecorations,
        },
      }),
      new Plugin({ props: { decorations: (state) => this.numberDecoration(state.doc) } }),
      keymap(baseKeymap),
    ]
  }

  /** O número, no começo do primeiro parágrafo, como o Word o escreve. */
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

  /** As opções do controle de alterações, quando ele está ligado. */
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

    // O corpo e o nó de fora andam juntos; se um passo anterior não chegou lá
    // fora, os passos deste cairiam com o deslocamento errado — no texto, logo
    // depois da referência. O corpo volta ao que o documento tem e a edição
    // fica de fora: perder uma tecla é melhor que escrever no lugar errado.
    const at = this.getPos()
    const current = at === undefined ? null : this.outer.state.doc.nodeAt(at)
    if (at === undefined || current === null || !current.content.eq(this.view.state.doc.content)) {
      if (current !== null) {
        const { doc } = this.view.state
        this.view.dispatch(
          this.view.state.tr.replace(0, doc.content.size, current.slice(0)).setMeta(FROM_OUTSIDE, true),
        )
      }
      return
    }

    // O controle de alterações é feito aqui, e não pelo `dispatchTransaction`
    // de fora: é aqui que a seleção do corpo existe, e é ela que o Backspace
    // controlado precisa deixar antes do que ficou excluído.
    const options = this.tracking()
    let effective = tr
    if (options !== null && shouldTrack(tr)) {
      try {
        effective = trackTransaction(tr, this.view.state, options.author(), new Date(), {
          composing: this.view.composing,
        })
      } catch (error) {
        // Melhor a edição sem controle que a edição perdida.
        console.error(error)
      }
    }
    const { state, transactions } = this.view.state.applyTransaction(effective)
    this.view.updateState(state)
    changed()

    const outer = this.outer.state.tr
    const offset = StepMap.offset(at + 1)
    for (const transaction of transactions) {
      for (const step of transaction.steps) {
        const mapped = step.map(offset)
        if (mapped !== null) outer.step(mapped)
      }
    }
    if (!outer.docChanged) return
    if (options !== null) outer.setMeta(SKIP_TRACKING, true)
    if (this.separateHistory) {
      this.separateHistory = false
      closeHistory(outer)
    }
    if (tr.getMeta('addToHistory') === false) outer.setMeta('addToHistory', false)
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

  /** O "editável" de fora mudou (somente leitura, modo de leitura). */
  refresh(): void {
    this.view.setProps({})
  }

  /**
   * Da referência para a nota: o cursor no fim do corpo. O corpo que ainda está
   * no depósito (a nota acabou de nascer) já recebe o foco — o que se digita
   * antes de a paginação o pôr na folha não vai parar no texto —, e a folha o
   * mostra quando ele chegar lá.
   */
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

  /** Seleciona um trecho do corpo — a ocorrência da busca. */
  select(from: number, to: number): void {
    const { state } = this.view
    const size = state.doc.content.size
    const selection = TextSelection.create(state.doc, Math.min(from, size), Math.min(to, size))
    this.view.dispatch(state.tr.setSelection(selection).setMeta(FROM_OUTSIDE, true))
    this.body.scrollIntoView({ block: 'nearest' })
  }

  /** Da nota para a referência: o cursor logo depois dela, no texto. */
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
    // A nota com o cursor saiu do documento (o desfazer que a inseriu, o
    // Backspace na referência): o foco volta ao texto, onde a seleção de fora
    // já está — sem isso ele ficava em lugar nenhum e o refazer não chegava.
    // No microtask: agora o editor de fora ainda está no meio da atualização.
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

/** A vista da referência, para `addNodeView`. */
export function noteRefView(editor: Editor) {
  return (props: NodeViewRendererProps): NoteBody => new NoteBody(props, editor)
}
