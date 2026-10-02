import { Node } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode, type Schema } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { DocumentNotes } from '@services/document/model.js'
import { NoteKind, noteLabel, noteLabels } from '@services/document/notes.js'
import { noteRefView } from './note-view.js'

/**
 * Notas de rodapé e de fim (M11): a referência como nó, com o corpo dentro.
 *
 * Um nó em linha e atômico, mas **com conteúdo**: os blocos da nota (`block+`).
 * Atômico porque o corpo não se edita no texto — ele tem um editor próprio, no
 * pé da página (`note-view.ts`) —, e com conteúdo porque é assim que ele viaja: copiar a
 * referência copia a nota, apagá-la apaga a nota, e o arquivo recebe de volta o
 * corpo que leu. O `nid` é o `w:id` que casa a referência com a nota no arquivo;
 * a colada ao lado da original perde o dela, e a gravação lhe dá uma nota própria.
 *
 * O número não é atributo: é a ordem da referência no documento, contada por
 * `noteLabels` e desenhada por decoração (`data-note-number`), que o JSON nunca
 * vê. A marca própria (`w:customMarkFollows`) é texto do nó, e não conta.
 */

export interface NoteRefOptions {
  /** A numeração do documento — ver `DocumentModel.notes`. Consultada a cada conta. */
  readonly notes: (() => DocumentNotes | undefined) | undefined
}

export const noteRefKey = new PluginKey<DecorationSet>('noteRef')

/** As referências do documento, em ordem — sem descer no corpo de nenhuma. */
export function noteRefsOf(doc: ProseMirrorNode): Array<{ node: ProseMirrorNode; pos: number }> {
  const found: Array<{ node: ProseMirrorNode; pos: number }> = []
  doc.descendants((node, pos) => {
    if (node.type.name !== 'noteRef') return true
    found.push({ node, pos })
    return false
  })
  return found
}

function markOf(node: ProseMirrorNode): string | null {
  const mark: unknown = node.attrs['mark']
  return typeof mark === 'string' && mark !== '' ? mark : null
}

/**
 * Se o corpo da nota leva o número desenhado no começo. A de marca própria
 * (`customMarkFollows`) não: o Word grava a marca no próprio corpo, e desenhá-la
 * de novo daria "**".
 */
export function drawsNoteNumber(node: ProseMirrorNode): boolean {
  return markOf(node) === null
}

/** O rótulo de cada referência do documento, na ordem do texto. */
export function noteRefLabels(doc: ProseMirrorNode, notes?: DocumentNotes): string[] {
  return noteLabels(
    noteRefsOf(doc).map(({ node }) => ({ kind: String(node.attrs['kind']), mark: markOf(node) })),
    notes,
  )
}

function numberDecorations(doc: ProseMirrorNode, notes: DocumentNotes | undefined): DecorationSet {
  const refs = noteRefsOf(doc)
  if (refs.length === 0) return DecorationSet.empty
  const labels = noteRefLabels(doc, notes)
  return DecorationSet.create(
    doc,
    refs.map(({ node, pos }, index) =>
      Decoration.node(
        pos,
        pos + node.nodeSize,
        markOf(node) === null ? { 'data-note-number': labels[index]! } : {},
        // O corpo (note-view.ts) lê daqui o número que desenha no começo da nota.
        { noteLabel: labels[index]! },
      ),
    ),
  )
}

/**
 * A colagem que repetiria uma nota do documento leva a nota sem `nid`: o
 * arquivo não aceita duas referências à mesma nota, e a gravação dá à colada uma
 * nota própria, com o mesmo corpo. Arrastar não é colar — a referência só muda
 * de lugar, e leva o `nid` junto.
 */
export function withoutRepeatedNotes(slice: Slice, doc: ProseMirrorNode, moving = false): Slice {
  if (moving) return slice
  const present = new Set(
    noteRefsOf(doc).map(({ node }) => `${String(node.attrs['kind'])}:${String(node.attrs['nid'])}`),
  )
  let changed = false
  const strip = (fragment: Fragment): Fragment => {
    const children: ProseMirrorNode[] = []
    fragment.forEach((child) => {
      if (child.type.name === 'noteRef' && child.attrs['nid'] !== null) {
        if (present.has(`${String(child.attrs['kind'])}:${String(child.attrs['nid'])}`)) {
          changed = true
          children.push(child.type.create({ ...child.attrs, nid: null }, child.content, child.marks))
          return
        }
      }
      children.push(child.isLeaf ? child : child.copy(strip(child.content)))
    })
    return Fragment.from(children)
  }
  const content = strip(slice.content)
  return changed ? new Slice(content, slice.openStart, slice.openEnd) : slice
}

function contentFromJson(element: HTMLElement, schema: Schema): Fragment {
  try {
    const raw = element.getAttribute('data-note-body')
    if (raw !== null) return Fragment.fromJSON(schema, JSON.parse(raw) as unknown)
  } catch {
    // Corpo ilegível: a nota chega vazia em vez de derrubar a colagem.
  }
  return Fragment.from(schema.nodes['paragraph']!.create())
}

export const NoteRef = Node.create<NoteRefOptions>({
  name: 'noteRef',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  content: 'block+',

  addOptions() {
    return { notes: undefined }
  },

  addAttributes() {
    return {
      kind: {
        default: NoteKind.Footnote,
        parseHTML: (element) =>
          element.getAttribute('data-kind') === NoteKind.Endnote ? NoteKind.Endnote : NoteKind.Footnote,
        renderHTML: (attributes) => ({ 'data-kind': String(attributes['kind']) }),
      },
      nid: {
        default: null,
        parseHTML: (element) => element.getAttribute('data-nid'),
        renderHTML: (attributes) =>
          typeof attributes['nid'] === 'string' ? { 'data-nid': attributes['nid'] } : {},
      },
      mark: {
        default: null,
        parseHTML: (element) => element.getAttribute('data-mark'),
        renderHTML: (attributes) =>
          typeof attributes['mark'] === 'string' ? { 'data-mark': attributes['mark'] } : {},
      },
    }
  },

  parseHTML() {
    return [
      {
        tag: 'sup[data-note-ref]',
        // Acima da marca de sobrescrito, que também reconhece o `<sup>` e, com a
        // mesma prioridade, venceria: a colagem perdia a referência (e a nota).
        priority: 100,
        // O corpo vai num atributo, e não como filhos: `<p>` dentro de `<p>` faz o
        // analisador de HTML fechar o parágrafo de fora, e a colagem partiria o
        // parágrafo em volta da referência.
        getContent: (element, schema) => contentFromJson(element as HTMLElement, schema),
      },
    ]
  },

  renderHTML({ node, HTMLAttributes }) {
    return [
      'sup',
      {
        ...HTMLAttributes,
        'data-note-ref': '',
        class: 'note-ref',
        'data-note-body': JSON.stringify(node.content.toJSON()),
      },
      markOf(node) ?? '',
    ]
  },

  addNodeView() {
    return noteRefView(this.editor)
  },

  renderText({ node }) {
    return markOf(node) ?? ''
  },

  addProseMirrorPlugins() {
    const notes = (): DocumentNotes | undefined => this.options.notes?.()
    return [
      new Plugin<DecorationSet>({
        key: noteRefKey,
        state: {
          init: (_config, state) => numberDecorations(state.doc, notes()),
          apply: (transaction, previous, _old, state) =>
            transaction.docChanged ? numberDecorations(state.doc, notes()) : previous,
        },
        props: {
          decorations: (state) => noteRefKey.getState(state),
          transformPasted: (slice, view) =>
            withoutRepeatedNotes(slice, view.state.doc, view.dragging?.move === true),
        },
      }),
    ]
  },
})

/**
 * `textBetween` sem o corpo das notas.
 *
 * O ProseMirror desce em todo nó que não é folha, e a referência de nota tem o
 * corpo dentro: o título com uma nota ia para o sumário, para a referência
 * cruzada e para o painel de navegação com o texto da nota colado nele.
 */
export function textBetweenWithoutNotes(
  node: ProseMirrorNode,
  from: number,
  to: number,
  blockSeparator?: string,
  leafText: (leaf: ProseMirrorNode) => string = () => '',
): string {
  let text = ''
  let first = true
  node.nodesBetween(from, to, (child, pos) => {
    if (child.type.name === 'noteRef') return false
    const own = child.isText
      ? (child.text ?? '').slice(Math.max(from, pos) - pos, to - pos)
      : child.isLeaf
        ? leafText(child)
        : ''
    if (
      child.isBlock &&
      ((child.isLeaf && own !== '') || child.isTextblock) &&
      blockSeparator !== undefined
    ) {
      if (first) first = false
      else text += blockSeparator
    }
    text += own
    return true
  })
  return text
}

/** A numeração do documento como o editor a conhece — para o papel, que não vê decorações. */
export function notesSetupOf(
  extensions: ReadonlyArray<{ name: string; options: unknown }>,
): DocumentNotes | undefined {
  const extension = extensions.find((candidate) => candidate.name === 'noteRef')
  return (extension?.options as NoteRefOptions | undefined)?.notes?.()
}

/**
 * Escreve o número nas referências do HTML do papel, na ordem em que aparecem.
 * `counters` atravessa as folhas: a conta é do documento, não da página.
 */
export function numberNotesForPrint(
  holder: HTMLElement,
  counters: Map<string, number>,
  notes: DocumentNotes | undefined,
): void {
  for (const element of holder.querySelectorAll<HTMLElement>('sup[data-note-ref]')) {
    element.removeAttribute('data-note-body')
    if (element.hasAttribute('data-mark')) continue
    const kind = element.getAttribute('data-kind') ?? NoteKind.Footnote
    const ordinal = counters.get(kind) ?? 0
    counters.set(kind, ordinal + 1)
    element.textContent = noteLabel(kind, ordinal, notes)
  }
}
