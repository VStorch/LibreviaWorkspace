import { Extension } from '@tiptap/core'
import { Fragment, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import {
  LIST_LEVELS,
  LIST_TYPES,
  itemDrawAttrs,
  listDrawAttrs,
  numberLists,
  parseNumbering,
  type ListNumbering as ListCount,
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
