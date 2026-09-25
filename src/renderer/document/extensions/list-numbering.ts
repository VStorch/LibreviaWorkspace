import { Extension, type CommandProps } from '@tiptap/core'
import { Fragment, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import {
  LIST_LEVELS,
  LIST_TYPES,
  itemDrawAttrs,
  listDrawAttrs,
  numberLists,
  parseNumbering,
  type ListInfo,
  type ListNumbering as ListCount,
  type LevelDef,
  type NumberingDef,
  type ListTreeReader,
} from '@services/document/list-numbering.js'

/**
 * A numeração das listas, desenhada como o Word a conta.
 *
 * A conta é de `list-numbering.ts`; aqui ficam os dois jeitos de entregá-la. Na
 * tela, por **decoração** — a marca é aparência, e gravá-la no nó faria cada item
 * parecer editado a cada lista que se mexe acima dele, e a gravação cirúrgica
 * reescreveria o documento. No papel, pelo atributo `listDraw`, que só existe na
 * cópia serializada para o PDF (`drawListsForPrint`): o serializador do
 * ProseMirror não vê decorações. Os dois produzem os mesmos atributos no mesmo
 * elemento, e a regra de `content-styles.ts` desenha os dois igual.
 */

export const PM_LIST_READER: ListTreeReader<ProseMirrorNode> = {
  typeOf: (node) => node.type.name,
  attrsOf: (node) => node.attrs,
  childrenOf: (node) => node.children,
}

export const listNumberingKey = new PluginKey<DecorationSet>('listNumbering')

/** Decorações de uma conta, na mesma pré-ordem em que ela foi feita. */
function decorationsOf(doc: ProseMirrorNode, numbering: ListCount): DecorationSet {
  if (numbering.lists.length === 0) return DecorationSet.empty
  const decorations: Decoration[] = []
  let list = 0
  let item = 0

  // `descendants` é pré-ordem, como `numberLists`: a n-ésima lista de um é a
  // n-ésima do outro. O item só conta quando é filho de lista — é a mesma regra
  // de lá.
  const walk = (node: ProseMirrorNode, pos: number, parentIsList: boolean): void => {
    const isList = LIST_TYPES.includes(node.type.name)
    if (isList) {
      const info = numbering.lists[list++]
      if (info !== undefined) decorations.push(Decoration.node(pos, pos + node.nodeSize, listDrawAttrs(info)))
    } else if (node.type.name === 'listItem' && parentIsList) {
      const label = numbering.labels[item++]
      if (label !== undefined)
        decorations.push(Decoration.node(pos, pos + node.nodeSize, itemDrawAttrs(label)))
    }
    node.forEach((child, offset) => walk(child, pos + 1 + offset, isList))
  }
  doc.forEach((child, offset) => walk(child, offset, false))
  return DecorationSet.create(doc, decorations)
}

/**
 * Os blocos do documento com a numeração gravada em `listDraw`, para o papel.
 *
 * Reconstrói só o que tem lista dentro; o resto volta como está — o recorte das
 * folhas compara blocos por posição, não por identidade, mas não há por que
 * copiar um parágrafo que não muda.
 */
export function drawListsForPrint(doc: ProseMirrorNode): ProseMirrorNode[] {
  const numbering = numberLists(doc, PM_LIST_READER)
  let list = 0
  let item = 0

  const rebuild = (node: ProseMirrorNode, parentIsList: boolean): ProseMirrorNode => {
    const isList = LIST_TYPES.includes(node.type.name)
    let draw: Record<string, string> | null = null
    if (isList) {
      const info = numbering.lists[list++]
      if (info !== undefined) draw = listDrawAttrs(info)
    } else if (node.type.name === 'listItem' && parentIsList) {
      const label = numbering.labels[item++]
      if (label !== undefined) draw = itemDrawAttrs(label)
    }
    if (node.isLeaf || (!isList && draw === null && !hasList(node))) return node

    const children: ProseMirrorNode[] = []
    node.forEach((child) => children.push(rebuild(child, isList)))
    const attrs = draw === null ? node.attrs : { ...node.attrs, listDraw: draw }
    return node.type.create(attrs, Fragment.fromArray(children), node.marks)
  }

  const blocks: ProseMirrorNode[] = []
  doc.forEach((child) => blocks.push(rebuild(child, false)))
  return blocks
}

function hasList(node: ProseMirrorNode): boolean {
  let found = false
  node.descendants((child) => {
    if (found) return false
    if (LIST_TYPES.includes(child.type.name)) found = true
    return !found
  })
  return found
}

/** Quantas listas envolvem a seleção — o nível do item, a contar de 1. */
export function listDepthAt(node: { depth: number; node: (depth: number) => ProseMirrorNode }): number {
  let depth = 0
  for (let level = node.depth; level > 0; level--) {
    if (LIST_TYPES.includes(node.node(level).type.name)) depth++
  }
  return depth
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    listNumbering: {
      /** "Reiniciar em 1" e "Definir valor inicial": a lista passa a contar à parte, a partir de `start`. */
      restartListNumbering: (start?: number) => ReturnType
      /** "Continuar numeração": a lista passa a contar com a anterior do mesmo tipo. */
      continueListNumbering: () => ReturnType
      /** Troca os níveis da lista em que está o cursor — ou cria a lista com eles. */
      applyListLevels: (kind: string, levels: readonly LevelDef[]) => ReturnType
    }
  }
}

/** Uma lista do documento, com o que a conta decidiu sobre ela. */
export interface ListEntry {
  readonly pos: number
  readonly node: ProseMirrorNode
  readonly info: ListInfo
}

/** Todas as listas, em pré-ordem, com a posição de cada uma. */
export function listEntries(doc: ProseMirrorNode): ListEntry[] {
  const { lists } = numberLists(doc, PM_LIST_READER)
  const entries: ListEntry[] = []
  doc.descendants((node, pos) => {
    if (LIST_TYPES.includes(node.type.name)) {
      const info = lists[entries.length]
      if (info !== undefined) entries.push({ pos, node, info })
    }
    return true
  })
  return entries
}

/** A lista mais de dentro em volta do cursor. */
function innermostList(state: EditorState): { pos: number; depth: number } | null {
  const { $from } = state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    if (LIST_TYPES.includes($from.node(depth).type.name)) return { pos: $from.before(depth), depth }
  }
  return null
}

/**
 * Uma chave de contagem que nenhuma outra lista tem.
 *
 * Só até a gravação: o sidecar cria o `w:num`, e na volta a chave passa a ser a
 * dele (`n12`). Até lá, é o que junta as listas que a pessoa disse que são uma.
 */
function freshKey(): string {
  return `nova-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}

/**
 * Dá à lista `target` e às sublistas que contam com ela uma nova definição.
 *
 * As sublistas vão junto porque herdam o `numId` da de fora: deixadas como
 * estavam, continuariam apontando a numeração antiga e seriam gravadas nela.
 */
function assignDefinition(
  tr: Transaction,
  target: ListEntry,
  oldKey: string,
  def: NumberingDef,
  numId: number | null,
): void {
  for (const entry of listEntries(tr.doc)) {
    const inside = entry.pos >= target.pos && entry.pos < target.pos + target.node.nodeSize
    if (!inside || entry.info.key !== oldKey) continue
    tr.setNodeMarkup(entry.pos, undefined, { ...entry.node.attrs, numId, numbering: def })
  }
}

/** A lista que começa em `pos` no documento da transação. */
function entryAt(tr: Transaction, pos: number): ListEntry | null {
  return listEntries(tr.doc).find((entry) => entry.pos === pos) ?? null
}

/**
 * Parte a lista antes do item do cursor, para que o que se muda valha dali para
 * baixo — é o que o Word faz com "Reiniciar em 1" no meio de uma lista.
 */
function splitAtCursor(state: EditorState, tr: Transaction, list: { pos: number; depth: number }): number {
  const { $from } = state.selection
  const index = $from.index(list.depth)
  if (index === 0) return list.pos
  const before = $from.before(list.depth + 1)
  tr.split(before, 1)
  return before + 1
}

const restart =
  (start = 1) =>
  ({ state, tr, dispatch }: CommandProps): boolean => {
    const list = innermostList(state)
    if (list === null || !Number.isInteger(start) || start < 0) return false
    if (dispatch === undefined) return true

    const pos = splitAtCursor(state, tr, list)
    const target = entryAt(tr, pos)
    if (target === null) return false

    const def: NumberingDef = {
      ...target.info.def,
      key: freshKey(),
      overrides: { [String(target.info.level)]: start },
    }
    assignDefinition(tr, target, target.info.key, def, null)
    return true
  }

const continuePrevious =
  () =>
  ({ state, tr, dispatch }: CommandProps): boolean => {
    const list = innermostList(state)
    if (list === null) return false
    const entries = listEntries(state.doc)
    const current = entries.find((entry) => entry.pos === list.pos)
    if (current === undefined) return false

    // A anterior do mesmo tipo, fora desta e sem ser uma que a contém — de
    // preferência no mesmo nível, que é o que o Word procura.
    const candidates = entries.filter(
      (entry) =>
        entry.pos < current.pos &&
        entry.pos + entry.node.nodeSize <= current.pos &&
        entry.info.kind === current.info.kind,
    )
    const previous =
      candidates.filter((entry) => entry.info.level === current.info.level).at(-1) ?? candidates.at(-1)
    if (previous === undefined || previous.info.key === current.info.key) return false
    if (dispatch === undefined) return true

    // A anterior que não tem chave própria (lista nova, ainda sem gravar) ganha
    // uma estável: a que a conta lhe dá vem da posição, e mudaria com a próxima
    // lista criada acima dela — separando de novo o que se acabou de juntar.
    let def = previous.info.def
    if (def.key.startsWith('nova') && !def.key.startsWith('nova-')) {
      def = { ...def, key: freshKey() }
      for (const entry of entries) {
        if (entry.info.key !== previous.info.key) continue
        tr.setNodeMarkup(entry.pos, undefined, { ...entry.node.attrs, numbering: def })
      }
    }

    const target = entryAt(tr, current.pos)
    if (target === null) return false
    assignDefinition(tr, target, current.info.key, def, previous.info.numId)
    return true
  }

const applyLevels =
  (kind: string, levels: readonly LevelDef[]) =>
  ({ state, tr, dispatch, commands }: CommandProps): boolean => {
    if (!LIST_TYPES.includes(kind)) return false
    let list = innermostList(state)
    if (list === null) {
      // Fora de lista: vira lista primeiro, com o comando de sempre, e os níveis
      // entram por cima. Um passo de desfazer para os dois.
      if (!(kind === 'bulletList' ? commands.toggleBulletList() : commands.toggleOrderedList())) return false
      list = innermostList(state.apply(tr))
      if (list === null) return false
    }
    if (dispatch === undefined) return true

    const target = entryAt(tr, list.pos)
    if (target === null) return false
    const def: NumberingDef = { key: freshKey(), levels: levels.map((level) => ({ ...level })) }

    for (const entry of listEntries(tr.doc)) {
      const inside = entry.pos >= target.pos && entry.pos < target.pos + target.node.nodeSize
      if (!inside || entry.info.key !== target.info.key) continue
      // O tipo do nó segue o nível: marcador é `bulletList`, número é
      // `orderedList` — é o que o leitor faria com a definição ao reabrir.
      const fmt = levels[entry.info.level]?.fmt ?? 'decimal'
      const type = state.schema.nodes[fmt === 'bullet' || fmt === 'none' ? 'bulletList' : 'orderedList']
      tr.setNodeMarkup(entry.pos, type, {
        ...entry.node.attrs,
        numId: null,
        numbering: def,
        // O que o arquivo disse do nível antigo não vale mais: a marca e o recuo
        // passam a ser os da definição nova.
        marker: null,
        indentMm: null,
        hangingMm: null,
      })
    }
    return true
  }

export const ListNumbering = Extension.create({
  name: 'listNumbering',
  // Antes do item de lista, que também quer o Tab.
  priority: 110,

  addGlobalAttributes() {
    return [
      {
        types: [...LIST_TYPES],
        attributes: {
          /**
           * A definição da numeração: os nove níveis, a chave da contagem e o
           * reinício. Ver `list-numbering.ts`.
           *
           * Em HTML vai como JSON, para que copiar e colar dentro do editor leve
           * a numeração junto — senão a lista colada voltava à padrão.
           */
          numbering: {
            default: null,
            parseHTML: (element) => {
              const raw = element.getAttribute('data-numbering')
              if (raw === null) return null
              try {
                return parseNumbering(JSON.parse(raw))
              } catch {
                return null
              }
            },
            renderHTML: (attributes) =>
              parseNumbering(attributes['numbering']) === null
                ? {}
                : { 'data-numbering': JSON.stringify(attributes['numbering']) },
          },
          /** O nível do arquivo, quando a árvore não o diz sozinha (`w:ilvl`). */
          level: {
            default: null,
            parseHTML: (element) => {
              const value = Number(element.getAttribute('data-level'))
              return element.hasAttribute('data-level') && Number.isInteger(value) ? value : null
            },
            renderHTML: (attributes) =>
              Number.isInteger(attributes['level']) ? { 'data-level': String(attributes['level']) } : {},
          },
        },
      },
      {
        types: [...LIST_TYPES, 'listItem'],
        attributes: {
          /**
           * O desenho calculado — só na cópia que vai para o papel.
           *
           * Nunca lido do HTML: colado de volta, viraria conteúdo do nó, e o item
           * pareceria editado para a gravação cirúrgica.
           */
          listDraw: {
            default: null,
            keepOnSplit: false,
            parseHTML: () => null,
            renderHTML: (attributes) => {
              const draw = attributes['listDraw'] as Record<string, string> | null
              return draw !== null && typeof draw === 'object' ? draw : {}
            },
          },
        },
      },
    ]
  },

  addCommands() {
    return {
      restartListNumbering: restart,
      continueListNumbering: continuePrevious,
      applyListLevels: applyLevels,
    }
  },

  addKeyboardShortcuts() {
    return {
      // O Word não passa do nono nível; aqui, a sublista a mais seria gravada no
      // nono e desenhada no décimo.
      Tab: ({ editor }) =>
        editor.isActive('listItem') && listDepthAt(editor.state.selection.$from) >= LIST_LEVELS,
    }
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: listNumberingKey,
        state: {
          init: (_config, state) => decorationsOf(state.doc, numberLists(state.doc, PM_LIST_READER)),
          apply(transaction, current, _old, state) {
            if (!transaction.docChanged) return current
            return decorationsOf(state.doc, numberLists(state.doc, PM_LIST_READER))
          },
        },
        props: {
          decorations: (state) => listNumberingKey.getState(state),
        },
      }),
    ]
  },
})
