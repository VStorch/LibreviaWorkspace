import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import { LINE_GAP_CLASS } from '../line-boxes.js'

/**
 * The block opening a sheet gets a margin the size of what was left of the previous one, plus the
 * margins and the gap between papers. As a **decoration**, not an element: a real spacer would
 * enter the selection, `Ctrl+A` and what gets copied.
 */
export const paginationKey = new PluginKey<DecorationSet>('pagination')

export const Pagination = Extension.create({
  name: 'pagination',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: paginationKey,
        state: {
          init: () => DecorationSet.empty,
          apply(transaction, current) {
            const next = transaction.getMeta(paginationKey) as DecorationSet | undefined
            if (next !== undefined) return next

            // Until the measurement arrives, decorations follow editing, or the sheets would
            // flicker.
            return current.map(transaction.mapping, transaction.doc)
          },
        },
        props: {
          decorations: (state) => paginationKey.getState(state),
        },
      }),
      // An anchored capture with text: its 1lh line (`content-styles.ts`) does not add up.
      new Plugin({
        props: {
          decorations: (state) => {
            const decorations: Decoration[] = []
            state.doc.forEach((node, offset) => {
              if (hasAnchoredImageAndText(node)) {
                decorations.push(Decoration.node(offset, offset + node.nodeSize, { [ANCHOR_TEXT_ATTR]: '' }))
              }
            })
            return DecorationSet.create(state.doc, decorations)
          },
        },
      }),
    ]
  },
})

/** Outside the history: undo brings back what the user wrote, not where the page fell. */
export function applyPageGaps(
  view: EditorView,
  written: ReadonlyMap<number, number>,
  gaps: ReadonlyMap<number, number>,
  { lines = new Map(), headers = [], columns = new Map() }: PageGapExtras = {},
): void {
  const decorations: Decoration[] = []

  // The repeated header is a copy in a widget: outside selection and editing.
  for (const header of headers) {
    if (header.position <= 0 || header.position > view.state.doc.content.size) continue
    decorations.push(
      Decoration.widget(header.position, () => repeatedHeader(header), {
        side: -1,
        ignoreSelection: true,
        key: `page-header:${header.position}:${header.height.toFixed(2)}:${header.html.length}`,
      }),
    )
  }

  // On a break between lines, the spacer is as wide as the line and comes before the first
  // character of the line opening the sheet: the line above ends where it already ended, with the
  // same justification. `vertical-align: top` does not add the descender to the gap.
  for (const [position, gap] of lines) {
    if (gap <= 0 || position <= 0 || position > view.state.doc.content.size) continue
    decorations.push(
      Decoration.widget(position, () => lineGap(gap), {
        side: -1,
        ignoreSelection: true,
        key: `page-line-gap:${gap.toFixed(2)}`,
      }),
    )
  }

  view.state.doc.descendants((node, offset) => {
    const dx = columns.get(offset)
    if (dx !== undefined && dx !== 0) {
      decorations.push(
        Decoration.node(offset, offset + node.nodeSize, { style: `transform:translateX(${dx}px)` }),
      )
    }

    // Column offsets can be negative; only zero is not written.
    const gap = written.get(offset)
    if (gap === undefined || (gap === 0 && !gaps.has(offset))) return

    decorations.push(
      Decoration.node(offset, offset + node.nodeSize, {
        style: `${node.type.name === 'tableCell' || node.type.name === 'tableHeader' ? 'padding-top' : 'margin-top'}:${gap}px`,
        'data-page-start': 'true',
        'data-page-shift': String(gaps.get(offset) ?? 0),
      }),
    )
  })

  view.dispatch(
    view.state.tr
      .setMeta(paginationKey, DecorationSet.create(view.state.doc, decorations))
      .setMeta('addToHistory', false),
  )
}

/** What pagination pushes beyond the gaps between blocks. */
export interface PageGapExtras {
  /** By the position of the line's first character. */
  readonly lines?: ReadonlyMap<number, number>
  readonly headers?: readonly RepeatedHeader[]
  /** A translation, not a margin: changing column does not change the height. */
  readonly columns?: ReadonlyMap<number, number>
}

/** A table's header rows, repeated at the top of a sheet. */
export interface RepeatedHeader {
  /** Start of the content of the first cell in the row opening the sheet. */
  readonly position: number
  /** With the original's columns and only the header rows. */
  readonly html: string
  readonly height: number
  /** From the corner of the cell content. */
  readonly offsetTop: number
  readonly offsetLeft: number
}

export const REPEATED_HEADER_CLASS = 'page-repeated-header'

function repeatedHeader(header: RepeatedHeader): HTMLElement {
  const element = document.createElement('div')
  element.className = REPEATED_HEADER_CLASS
  element.contentEditable = 'false'
  element.setAttribute('aria-hidden', 'true')
  element.style.cssText =
    `position:absolute;margin-top:${-header.offsetTop}px;margin-left:${-header.offsetLeft}px;` +
    `height:${header.height}px;pointer-events:none;user-select:none;`
  element.innerHTML = header.html
  return element
}

/** `data-page-shift` is what measuring subtracts. */
function lineGap(gap: number): HTMLElement {
  const element = document.createElement('span')
  element.className = LINE_GAP_CLASS
  element.contentEditable = 'false'
  element.setAttribute('aria-hidden', 'true')
  element.dataset.pageShift = String(gap)
  element.style.cssText = `display:inline-block;width:100%;height:${gap}px;vertical-align:top;line-height:0;`
  return element
}

/** An anchored capture **and** text; see `content-styles.ts`. */
export const ANCHOR_TEXT_ATTR = 'data-anchor-text'

export function hasAnchoredImageAndText(node: {
  readonly isTextblock: boolean
  readonly textContent: string
  forEach: (callback: (child: { type: { name: string }; attrs: Record<string, unknown> }) => void) => void
}): boolean {
  if (!node.isTextblock || node.textContent.trim() === '') return false
  let anchored = false
  node.forEach((child) => {
    if (child.type.name === 'image' && child.attrs['anchored'] === true) anchored = true
  })
  return anchored
}

export function isPaginationOnly(transaction: {
  getMeta: (key: PluginKey<DecorationSet>) => unknown
}): boolean {
  return transaction.getMeta(paginationKey) !== undefined
}
