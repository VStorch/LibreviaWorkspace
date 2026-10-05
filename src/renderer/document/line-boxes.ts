import type { EditorView } from '@tiptap/pm/view'

/**
 * The paginator only breaks where it is told it can: without lines, a half-page paragraph would
 * move down whole. The measure comes from `Range.getClientRects()`, which covers the font area, not
 * the line box; the boundary between two lines is the midpoint between one's foot and the other's
 * top.
 */
export interface ParagraphLines {
  /** CSS pixels from the top of the block border, **without** the page gaps applied inside it. */
  readonly starts: readonly number[]
  readonly shift: number
  /** Document position of the first character of the line `starts[index]` opens. */
  readonly positionOf: (index: number) => number | null
}

export const LINE_GAP_CLASS = 'page-line-gap'

interface Piece {
  readonly top: number
  readonly bottom: number
  /** Screen coordinates, to search for the character opening the line. */
  readonly clientTop: number
  readonly node: Node
  /** Above zero, the line starts in the middle of the text node. */
  readonly rectIndex: number
}

interface Line {
  top: number
  bottom: number
  readonly first: Piece
}

export function measureLines(view: EditorView, element: HTMLElement): ParagraphLines {
  const box = element.getBoundingClientRect()
  // With zoom (`transform`) the rectangle comes at screen scale; pagination counts in CSS pixels.
  const scale = element.offsetHeight > 0 && box.height > 0 ? box.height / element.offsetHeight : 1

  const pieces: Piece[] = []
  let shift = 0
  const push = (rect: DOMRect, node: Node, rectIndex: number): void => {
    if (rect.height <= 0) return
    pieces.push({
      top: (rect.top - box.top) / scale - shift,
      bottom: (rect.bottom - box.top) / scale - shift,
      clientTop: rect.top,
      node,
      rectIndex,
    })
  }

  const walker = document.createTreeWalker(element, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
    acceptNode: (node) => {
      if (!(node instanceof HTMLElement)) return NodeFilter.FILTER_ACCEPT
      // A gap from an earlier pass is a push, not a line.
      if (node.classList.contains(LINE_GAP_CLASS)) {
        // Hidden by measuring (`usePagination`): it pushes nothing.
        if (node.style.display !== 'none') shift += Number(node.dataset.pageShift ?? 0)
        return NodeFilter.FILTER_REJECT
      }
      if (node.classList.contains('ProseMirror-separator')) return NodeFilter.FILTER_REJECT
      const position = getComputedStyle(node).position
      if (position === 'absolute' || position === 'fixed') return NodeFilter.FILTER_REJECT
      if (node.tagName === 'IMG' || node.tagName === 'BR' || node.contentEditable === 'false') {
        push(node.getBoundingClientRect(), node, 0)
        return NodeFilter.FILTER_REJECT
      }
      return NodeFilter.FILTER_SKIP
    },
  })

  const range = document.createRange()
  for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
    if (!(node instanceof Text) || node.length === 0) continue
    range.selectNodeContents(node)
    Array.from(range.getClientRects()).forEach((rect, index) => push(rect, node, index))
  }

  // A new line when the **middle** of the piece passes the current one's foot: with tight line
  // spacing the areas overlap, and a superscript has another top on the same line.
  const lines: Line[] = []
  for (const piece of pieces) {
    const current = lines.at(-1)
    const center = (piece.top + piece.bottom) / 2
    if (current === undefined || piece.rectIndex > 0 || center > current.bottom) {
      lines.push({ top: piece.top, bottom: piece.bottom, first: piece })
    } else {
      current.top = Math.min(current.top, piece.top)
      current.bottom = Math.max(current.bottom, piece.bottom)
    }
  }

  const starts: number[] = []
  for (let index = 1; index < lines.length; index++) {
    starts.push((lines[index - 1]!.bottom + lines[index]!.top) / 2)
  }

  return {
    starts,
    shift,
    positionOf: (index) => {
      const line = lines[index + 1]
      return line === undefined ? null : positionOfLine(view, line.first, range)
    },
  }
}

/**
 * Binary search in the text node, only for the chosen break: measuring happens every frame,
 * breaking is rare.
 */
function positionOfLine(view: EditorView, first: Piece, range: Range): number | null {
  try {
    if (!(first.node instanceof Text) || first.rectIndex === 0) {
      const parent = first.node.parentNode
      if (parent === null) return null
      if (first.node instanceof Text) return view.posAtDOM(first.node, 0)
      return view.posAtDOM(parent, Array.prototype.indexOf.call(parent.childNodes, first.node) as number)
    }

    const text = first.node
    range.selectNodeContents(text)
    const previousTop = range.getClientRects()[first.rectIndex - 1]?.top ?? first.clientTop - 1
    let low = 0
    let high = text.length - 1
    let found = text.length
    while (low <= high) {
      const middle = (low + high) >> 1
      range.setStart(text, middle)
      range.setEnd(text, middle + 1)
      const rects = range.getClientRects()
      const rect = rects[rects.length - 1]
      // Leftover space at the end hangs on the line above.
      if (rect !== undefined && rect.top > previousTop + 1) {
        found = middle
        high = middle - 1
      } else {
        low = middle + 1
      }
    }
    return view.posAtDOM(text, found)
  } catch {
    return null
  }
}
