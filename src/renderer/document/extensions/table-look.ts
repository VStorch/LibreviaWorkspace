import { Extension, type Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { TableMap } from '@tiptap/pm/tables'
import { Plugin } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import {
  cellBordersFromAttr,
  cellLookPatch,
  cellBordersToCss,
  resolvedColumnWidths,
  tableDraftFrom,
  type CellLook,
  type TableDraft,
} from '@services/document/table-format.js'
import { mmToPx, pxToMm, twipsToPx } from '@services/units.js'

/**
 * TableKit gives structure and no appearance. `applyTableDraft` writes the form in a single
 * transaction: four would take four `Ctrl+Z` and four document remeasurements.
 */

const CELL_TYPES: readonly string[] = ['tableCell', 'tableHeader']

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    tableLook: {
      /** With a function, each cell gets what it returns from what it has. */
      setCellLook: (look: CellLook | ((current: CellLook) => CellLook)) => ReturnType
      /** The whole grid: the writer drops a partial `w:tblGrid`. See `resolvedColumnWidths`. */
      setColumnWidths: (widths: readonly number[]) => ReturnType
    }
  }
}

export const TableLook = Extension.create({
  name: 'tableLook',

  /**
   * On screen, `TableView` builds its own `<table>` without the attributes: the decoration lands on
   * its wrapper.
   */
  addProseMirrorPlugins() {
    return [
      new Plugin({
        // Redone only when the document changes, not on every selection or pagination.
        state: {
          init: (_config, state) => cellMarginDecorations(state.doc),
          apply: (transaction, current) =>
            transaction.docChanged ? cellMarginDecorations(transaction.doc) : current,
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
        },
      }),
    ]
  },

  addGlobalAttributes() {
    return [
      {
        types: ['table'],
        attributes: {
          /** In twips, as the reader resolved it; without it, the editor template's margin. */
          cellMargins: {
            default: null,
            parseHTML: () => null,
            renderHTML: (attributes: Record<string, unknown>) => {
              const css = cellMarginsCss(attributes['cellMargins'])
              return css === null ? {} : { style: `--cell-margins: ${css}` }
            },
          },
        },
      },
      {
        types: [...CELL_TYPES],
        attributes: {
          /** Canonical text; see `table-format.ts`. */
          borders: {
            default: null,
            // No `parseHTML`: a table pasted from outside does not get borders its source did not
            // declare.
            parseHTML: () => null,
            renderHTML: (attributes: Record<string, unknown>) => {
              const css = cellBordersToCss(cellBordersFromAttr(attributes['borders']))
              return css === '' ? {} : { style: css }
            },
          },

          shading: {
            default: null,
            parseHTML: () => null,
            renderHTML: (attributes: Record<string, unknown>) => {
              const fill = attributes['shading']
              return typeof fill === 'string' && fill !== '' ? { style: `background-color:${fill}` } : {}
            },
          },
        },
      },
    ]
  },

  addCommands() {
    return {
      setCellLook:
        (look) =>
        ({ state, tr, dispatch }) => {
          let touched = false

          for (const range of state.selection.ranges) {
            state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, pos) => {
              if (!CELL_TYPES.includes(node.type.name)) return true
              const current: CellLook = {
                borders: typeof node.attrs['borders'] === 'string' ? node.attrs['borders'] : null,
                shading: typeof node.attrs['shading'] === 'string' ? node.attrs['shading'] : null,
              }
              const next = typeof look === 'function' ? look(current) : look

              // Writing the same value again is still an edit, and the whole table would be
              // rewritten.
              if (next.borders !== current.borders) tr.setNodeAttribute(pos, 'borders', next.borders)
              if (next.shading !== current.shading) tr.setNodeAttribute(pos, 'shading', next.shading)
              touched = true
              return false
            })
          }

          if (!touched) return false
          if (dispatch !== undefined) dispatch(tr)
          return true
        },

      setColumnWidths:
        (widths) =>
        ({ state, tr, dispatch }) => {
          const found = tableAt(state.selection.$from)
          if (found === null) return false

          const map = TableMap.get(found.node)
          if (map.width !== widths.length) return false

          // By the **grid**: a merged cell appears several times, and writing it per column would
          // overwrite the measure.
          const written = new Set<number>()

          for (let column = 0; column < map.width; column += 1) {
            for (let row = 0; row < map.height; row += 1) {
              const relative = map.map[row * map.width + column]
              if (relative === undefined || written.has(relative)) continue
              written.add(relative)

              const cell = found.node.nodeAt(relative)
              if (cell === null) continue

              const rect = map.findCell(relative)
              const own = widths.slice(rect.left, rect.right)
              tr.setNodeMarkup(found.start + relative, undefined, { ...cell.attrs, colwidth: own })
            }
          }

          if (dispatch !== undefined) dispatch(tr)
          return true
        },
    }
  },
})

interface FoundTable {
  readonly node: ProseMirrorNode
  readonly start: number
}

function tableAt($pos: {
  depth: number
  node: (depth: number) => ProseMirrorNode
  start: (depth: number) => number
}): FoundTable | null {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const node = $pos.node(depth)
    if (node.type.spec['tableRole'] === 'table') return { node, start: $pos.start(depth) }
  }

  return null
}

/** `null` outside a table: greys out the menu item. */
export interface TablePlacement {
  readonly draft: TableDraft
  /** The resolved grid, including columns the document does not declare. */
  readonly columnWidths: readonly number[]
  readonly column: number
}

export function tablePlacementAt(editor: Editor, contentWidthPx: number): TablePlacement | null {
  const { $from } = editor.state.selection
  const found = tableAt($from)
  if (found === null) return null

  const cell = cellAt($from)
  if (cell === null) return null

  const map = TableMap.get(found.node)
  const rect = map.findCell(cell.pos - found.start)

  // From the first row, where TableKit reads the `colgroup` from.
  const declared: (number | null)[] = []
  const first = found.node.firstChild
  if (first !== null) {
    first.forEach((child) => {
      const widths = child.attrs['colwidth']
      const span = Number(child.attrs['colspan']) || 1
      for (let index = 0; index < span; index += 1) {
        const width = Array.isArray(widths) ? Number(widths[index]) : Number.NaN
        declared.push(Number.isFinite(width) && width > 0 ? width : null)
      }
    })
  }

  const columnWidths = resolvedColumnWidths(declared, contentWidthPx)

  return {
    column: rect.left,
    columnWidths,
    draft: tableDraftFrom({
      borders: cell.node.attrs['borders'],
      shading: cell.node.attrs['shading'],
      columnWidthMm: Math.round(pxToMm(columnWidths[rect.left] ?? 0) * 10) / 10,
      headerRow: found.node.firstChild?.firstChild?.type.name === 'tableHeader',
    }),
  }
}

function cellAt($pos: {
  depth: number
  node: (depth: number) => ProseMirrorNode
  before: (depth: number) => number
}): { readonly node: ProseMirrorNode; readonly pos: number } | null {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const node = $pos.node(depth)
    if (CELL_TYPES.includes(node.type.name)) return { node, pos: $pos.before(depth) }
  }

  return null
}

/** The header row last: `toggleHeaderRow` changes the node types and invalidates positions. */
export function applyTableDraft(editor: Editor, draft: TableDraft, contentWidthPx: number): boolean {
  const placement = tablePlacementAt(editor, contentWidthPx)
  if (placement === null) return false

  const chain = editor.chain().focus()

  // Only what the user touched; see `cellLookPatch`.
  const opened = placement.draft
  chain.setCellLook((current) => cellLookPatch(current, opened, draft))

  // Rewriting the untouched grid would declare in pixels what the file measured in twips.
  if (draft.columnWidthMm !== null && draft.columnWidthMm !== opened.columnWidthMm) {
    const widths = [...placement.columnWidths]
    widths[placement.column] = Math.max(1, Math.round(mmToPx(draft.columnWidthMm)))
    chain.setColumnWidths(widths)
  }

  if (draft.headerRow !== placement.draft.headerRow) chain.toggleHeaderRow()

  return chain.run()
}

/** `"0 108 0 108"` (twips) as CSS `padding`; anything that is not four numbers is dropped. */
export function cellMarginsCss(value: unknown): string | null {
  if (typeof value !== 'string') return null
  const sides = value.trim().split(/\s+/).map(Number)
  if (sides.length !== 4 || sides.some((side) => !Number.isFinite(side) || side < 0)) return null
  return sides.map((side) => `${Math.round(twipsToPx(side) * 100) / 100}px`).join(' ')
}

/**
 * Without descending into paragraphs: walking the text on every edit would cost the whole document
 * per key press.
 */
function cellMarginDecorations(doc: ProseMirrorNode): DecorationSet {
  const decorations: Decoration[] = []
  doc.descendants((node, pos) => {
    if (node.isTextblock || node.isAtom) return false
    if (node.type.name !== 'table') return true
    const css = cellMarginsCss(node.attrs['cellMargins'])
    if (css !== null) {
      decorations.push(Decoration.node(pos, pos + node.nodeSize, { style: `--cell-margins: ${css}` }))
    }
    return true
  })
  return DecorationSet.create(doc, decorations)
}
