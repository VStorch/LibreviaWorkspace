import { Extension, type CommandProps } from '@tiptap/core'
import { Fragment, Slice, type Node as ProseMirrorNode, type Schema } from '@tiptap/pm/model'
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
 * Na tela, por **decoração**: gravar a marca no nó faria cada item parecer
 * editado quando uma lista acima muda. No papel, pelo atributo `listDraw`, só na
 * cópia serializada (`drawListsForPrint`), porque o serializador não vê
 * decorações. Os dois produzem os mesmos atributos.
 */

export const PM_LIST_READER: ListTreeReader<ProseMirrorNode> = {
  typeOf: (node) => node.type.name,
  attrsOf: (node) => node.attrs,
  childrenOf: (node) => node.children,
}

export const listNumberingKey = new PluginKey<DecorationSet>('listNumbering')

function decorationsOf(doc: ProseMirrorNode, numbering: ListCount): DecorationSet {
  if (numbering.lists.length === 0) return DecorationSet.empty
  const decorations: Decoration[] = []
  let list = 0
  let item = 0

  // Pré-ordem, como `numberLists`; o item só conta quando é filho de lista.
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

/** Só o que tem lista dentro é reconstruído. */
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

/**
 * A chave vem do documento de origem e pode existir aqui com outra definição.
 * Mesma chave com os mesmos níveis é cópia de dentro do documento, e continua,
 * como no Word.
 */
export function renamePastedKeys(slice: Slice, doc: ProseMirrorNode): Slice {
  const existing = new Map<string, string>()
  for (const entry of listEntries(doc)) {
    const own = parseNumbering(entry.node.attrs['numbering'])
    if (own !== null) existing.set(own.key, JSON.stringify(own.levels))
  }

  const renamed = new Map<string, string>()
  let changed = false
  const rename = (node: ProseMirrorNode): ProseMirrorNode => {
    const children: ProseMirrorNode[] = []
    node.forEach((child) => children.push(rename(child)))
    const own = LIST_TYPES.includes(node.type.name) ? parseNumbering(node.attrs['numbering']) : null
    let attrs = node.attrs
    if (own !== null && existing.get(own.key) !== JSON.stringify(own.levels)) {
      const key = renamed.get(own.key) ?? freshKey()
      renamed.set(own.key, key)
      attrs = { ...node.attrs, numbering: { ...own, key } }
      changed = true
    }
    if (
      node.isText ||
      (attrs === node.attrs && children.every((child, index) => child === node.maybeChild(index)))
    ) {
      return node
    }
    return node.type.create(attrs, Fragment.fromArray(children), node.marks)
  }

  const content: ProseMirrorNode[] = []
  slice.content.forEach((node) => content.push(rename(node)))
  return changed ? new Slice(Fragment.fromArray(content), slice.openStart, slice.openEnd) : slice
}

/** O nível do item, a contar de 1. */
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

export interface ListEntry {
  readonly pos: number
  readonly node: ProseMirrorNode
  readonly info: ListInfo
}

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

function innermostList(state: EditorState): { pos: number; depth: number } | null {
  const { $from } = state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    if (LIST_TYPES.includes($from.node(depth).type.name)) return { pos: $from.before(depth), depth }
  }
  return null
}

/** Só até a gravação: o sidecar cria o `w:num`, e a chave passa a ser a dele (`n12`). */
function freshKey(): string {
  return `nova-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}

/** As sublistas herdam o `numId` da de fora, e vão junto. */
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

function entryAt(tr: Transaction, pos: number): ListEntry | null {
  return listEntries(tr.doc).find((entry) => entry.pos === pos) ?? null
}

/** Como o Word faz com "Reiniciar em 1" no meio de uma lista. */
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

    // A anterior do mesmo tipo, de preferência no mesmo nível, como o Word procura.
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

    // A lista nova ainda sem gravar ganha chave estável: a da conta vem da posição e mudaria.
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
      // Fora de lista, vira lista primeiro: um passo de desfazer para os dois.
      if (!(kind === 'bulletList' ? commands.toggleBulletList() : commands.toggleOrderedList())) return false
      list = innermostList(state.apply(tr))
      if (list === null) return false
    }
    if (dispatch === undefined) return true

    const target = entryAt(tr, list.pos)
    if (target === null) return false
    relabelList(tr, state.schema, target, levels)
    return true
  }

/** O tipo do nó segue o nível, como o leitor faria ao reabrir. */
function relabelList(tr: Transaction, schema: Schema, target: ListEntry, levels: readonly LevelDef[]): void {
  // Os níveis que não mudaram são copiados da definição de origem na gravação.
  const abstractId = target.info.def.abstractId
  const def: NumberingDef = {
    key: freshKey(),
    ...(abstractId === undefined ? {} : { abstractId }),
    levels: levels.map((level) => ({ ...level })),
  }
  const end = target.pos + target.node.nodeSize
  for (const entry of listEntries(tr.doc)) {
    if (entry.pos < target.pos || entry.pos >= end || entry.info.key !== target.info.key) continue
    const fmt = levels[entry.info.level]?.fmt ?? 'decimal'
    const type = schema.nodes[fmt === 'bullet' || fmt === 'none' ? 'bulletList' : 'orderedList']
    tr.setNodeMarkup(entry.pos, type, {
      ...entry.node.attrs,
      numId: null,
      numbering: def,
      marker: null,
      indentMm: null,
      hangingMm: null,
    })
  }
}

export const ListNumbering = Extension.create({
  name: 'listNumbering',
  priority: 110,

  addGlobalAttributes() {
    return [
      {
        types: [...LIST_TYPES],
        attributes: {
          /** Os nove níveis, a chave e o reinício. Em HTML vai como JSON, para copiar e colar levar a numeração. */
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
          /** Só na cópia para o papel; nunca lido do HTML, senão o item pareceria editado. */
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
      // O Word não passa do nono nível.
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
          transformPasted: (slice, view) => renamePastedKeys(slice, view.state.doc),
        },
      }),
    ]
  },
})
