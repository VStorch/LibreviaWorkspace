import { Node } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode, type Schema } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { DocumentNotes } from '@services/document/model.js'
import { NoteKind, noteLabels } from '@services/document/notes.js'
import { sectionBreakIn, type SectionBlock } from '@services/document/sections.js'
import { noteRefView } from './note-view.js'

/**
 * Notas de rodapé e de fim: a referência como nó, com o corpo dentro.
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

export const noteRefKey = new PluginKey<NoteRefState>('noteRef')

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

/**
 * A posição da referência cuja nota contém `pos`, ou `null` fora de qualquer
 * nota. É por ela que o comando que achou um trecho dentro de uma nota sabe que
 * a seleção é a do corpo, que tem editor próprio (`note-view.ts`).
 */
export function noteRefAround(doc: ProseMirrorNode, pos: number): number | null {
  if (pos < 0 || pos > doc.content.size) return null
  const $pos = doc.resolve(pos)
  for (let depth = $pos.depth; depth > 0; depth--) {
    if ($pos.node(depth).type.name === 'noteRef') return $pos.before(depth)
  }
  return null
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

/**
 * A seção de cada referência, pela ordem de `noteRefsOf`: a contagem das marcas
 * de seção (`sectionBreak`) dos blocos de primeiro nível antes dela — a marca
 * fecha a seção, como o `w:sectPr` no parágrafo.
 */
function noteRefSections(doc: ProseMirrorNode): number[] {
  const sections: number[] = []
  let section = 0
  doc.forEach((block) => {
    block.descendants((node) => {
      if (node.type.name !== 'noteRef') return true
      sections.push(section)
      return false
    })
    if (sectionBreakIn(block as unknown as SectionBlock) !== null) section += 1
  })
  return sections
}

/**
 * O rótulo de cada referência do documento, na ordem do texto. `pages` é a folha
 * de cada uma (pelo índice), para a numeração que reinicia a cada página.
 */
export function noteRefLabels(
  doc: ProseMirrorNode,
  notes?: DocumentNotes,
  pages: readonly (number | undefined)[] = [],
): string[] {
  const sections = noteRefSections(doc)
  return noteLabels(
    noteRefsOf(doc).map(({ node }, index) => {
      const page = pages[index]
      return {
        kind: String(node.attrs['kind']),
        mark: markOf(node),
        section: sections[index] ?? 0,
        ...(page === undefined ? {} : { page }),
      }
    }),
    notes,
  )
}

interface NoteRefState {
  readonly decorations: DecorationSet
  /** O rótulo de cada referência, pela ordem de `noteRefsOf`. */
  readonly labels: readonly string[]
  /** A folha de cada referência, que a paginação conta — ver `setNotePages`. */
  readonly pages: readonly (number | undefined)[]
}

function noteRefState(
  doc: ProseMirrorNode,
  notes: DocumentNotes | undefined,
  pages: readonly (number | undefined)[],
): NoteRefState {
  const refs = noteRefsOf(doc)
  if (refs.length === 0) return { decorations: DecorationSet.empty, labels: [], pages }
  const labels = noteRefLabels(doc, notes, pages)
  const decorations = DecorationSet.create(
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
  return { decorations, labels, pages }
}

/** Os rótulos que a tela mostra agora, pela ordem de `noteRefsOf` — o papel e o `NOTEREF` usam os mesmos. */
export function noteLabelsOf(state: EditorState): readonly string[] {
  // Sem o plugin (um estado montado à parte), a conta padrão do documento.
  return noteRefKey.getState(state)?.labels ?? noteRefLabels(state.doc)
}

export function notePagesOf(state: EditorState): readonly (number | undefined)[] {
  return noteRefKey.getState(state)?.pages ?? []
}

/**
 * Dá à numeração a folha de cada referência (`numRestart` `eachPage`). Vem da
 * paginação, depois de ela assentar; a transação não muda o documento.
 */
export function setNotePages(tr: Transaction, pages: readonly (number | undefined)[]): Transaction {
  return tr.setMeta(noteRefKey, pages).setMeta('addToHistory', false)
}

/**
 * A folha de cada nota de rodapé, pelo índice da referência: a folha em que a
 * nota começa — e a paginação a põe na folha da referência.
 */
export function footnotePagesOf(
  areas: ReadonlyArray<{
    readonly sheet: number
    readonly kind: string
    readonly items: ReadonlyArray<{ readonly index: number; readonly fromLine: number }>
  }>,
): (number | undefined)[] {
  const pages: (number | undefined)[] = []
  for (const area of areas) {
    if (area.kind !== NoteKind.Footnote) continue
    for (const item of area.items) if (item.fromLine === 0) pages[item.index] = area.sheet
  }
  return pages
}

/** As duas listas de folhas dizem o mesmo. */
export function samePages(
  left: readonly (number | undefined)[],
  right: readonly (number | undefined)[],
): boolean {
  const length = Math.max(left.length, right.length)
  for (let index = 0; index < length; index++) if (left[index] !== right[index]) return false
  return true
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
      new Plugin<NoteRefState>({
        key: noteRefKey,
        state: {
          init: (_config, state) => noteRefState(state.doc, notes(), []),
          apply: (transaction, previous, _old, state) => {
            const pages = transaction.getMeta(noteRefKey) as readonly (number | undefined)[] | undefined
            if (!transaction.docChanged && pages === undefined) return previous
            // As folhas da paginação de antes valem até ela assentar de novo.
            return noteRefState(state.doc, notes(), pages ?? previous.pages)
          },
        },
        props: {
          decorations: (state) => noteRefKey.getState(state)?.decorations ?? null,
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

/**
 * Escreve o número nas referências do HTML do papel, na ordem em que aparecem:
 * `labels` são os rótulos das referências desta folha, na mesma ordem (ver
 * `print-source.ts`), os mesmos da tela — reinícios por folha e seção incluídos.
 */
export function numberNotesForPrint(holder: HTMLElement, labels: readonly string[]): void {
  let index = 0
  for (const element of holder.querySelectorAll<HTMLElement>('sup[data-note-ref]')) {
    element.removeAttribute('data-note-body')
    const label = labels[index]
    index += 1
    if (element.hasAttribute('data-mark') || label === undefined) continue
    element.textContent = label
  }
}
