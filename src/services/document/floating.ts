import { pageDimensionsMm, type DocumentNode, type PageSetup } from './model.js'
import { bandForPage, pageLabel } from './band.js'

/** Mirrors the sidecar's `FloatDto`, in millimetres: screen and paper convert once each. */
export interface FloatingObject {
  /** `rule` is the line under a corporate header: a flat shape with an outline and no content. */
  readonly kind: 'image' | 'text' | 'rule'
  readonly src?: string | undefined
  readonly content?: DocumentNode[] | undefined
  readonly widthMm: number
  readonly heightMm: number
  /** Degrees, clockwise. */
  readonly rotation: number
  readonly hFrom: string
  readonly hOffsetMm?: number | undefined
  readonly hAlign?: string | undefined
  readonly vFrom: string
  readonly vOffsetMm?: number | undefined
  readonly vAlign?: string | undefined
  readonly behind: boolean
  readonly wrap: string
  /** The piece's position inside the shape group, added after resolving the anchor. */
  readonly dxMm?: number | undefined
  readonly dyMm?: number | undefined
  /**
   * Band objects only; the box is regenerated whole, because typing opens and closes paragraphs.
   */
  readonly bid?: string | undefined
  /** Solid color and stroke only; the rest is not drawn and goes to the inventory. */
  readonly fill?: string | undefined
  readonly line?: string | undefined
  readonly lineWidthPt?: number | undefined
  readonly dash?: boolean | undefined
}

/** So screen and paper draw the same. A zero-width stroke means none. */
export function frameOf(object: FloatingObject): { background?: string; border?: string } {
  const frame: { background?: string; border?: string } = {}

  if (typeof object.fill === 'string' && object.fill.length > 0) frame.background = object.fill

  const width = object.lineWidthPt ?? 0
  if (typeof object.line === 'string' && object.line.length > 0 && width > 0) {
    frame.border = `${width}pt ${object.dash === true ? 'dashed' : 'solid'} ${object.line}`
  }

  return frame
}

/** In millimetres from the sheet edge. */
export interface FloatingBox {
  readonly leftMm: number
  readonly topMm: number
  readonly widthMm: number
  readonly heightMm: number
  readonly rotation: number
  readonly behind: boolean
}

/**
 * The most common vertical origin is the paragraph, which only has a position after pagination, so
 * `anchorTopMm` is a parameter. Rotation passes through as is, because Word positions the box
 * unrotated and rotates around the center, like `transform: rotate()`.
 */
export function placeFloating(object: FloatingObject, page: PageSetup, anchorTopMm: number): FloatingBox {
  const { width, height } = pageDimensionsMm(page)
  const columnLeft = page.margins.left
  const columnWidth = width - page.margins.left - page.margins.right

  const leftMm = (() => {
    // Alignment beats offset: OOXML carries one or the other, never both, and with an alignment
    // there is no offset.
    if (object.hAlign !== undefined) {
      const box = referenceH(object.hFrom, page, width, columnLeft, columnWidth)
      if (object.hAlign === 'center') return box.start + (box.size - object.widthMm) / 2
      if (object.hAlign === 'right') return box.start + box.size - object.widthMm
      return box.start
    }

    const box = referenceH(object.hFrom, page, width, columnLeft, columnWidth)
    return box.start + (object.hOffsetMm ?? 0)
  })()

  const topMm = (() => {
    const offset = object.vOffsetMm ?? 0
    switch (object.vFrom) {
      case 'page':
        return offset
      case 'topMargin':
        return offset
      case 'bottomMargin':
        return height - page.margins.bottom + offset
      case 'margin':
        return page.margins.top + offset
      // `paragraph` and `line` are the same to us: the exact line inside the paragraph would
      // require measuring each line, and the difference is one line height.
      default:
        return anchorTopMm + offset
    }
  })()

  return {
    leftMm: leftMm + (object.dxMm ?? 0),
    topMm: topMm + (object.dyMm ?? 0),
    widthMm: object.widthMm,
    heightMm: object.heightMm,
    rotation: object.rotation,
    behind: object.behind,
  }
}

/** The horizontal band the offset refers to. */
function referenceH(
  from: string,
  page: PageSetup,
  width: number,
  columnLeft: number,
  columnWidth: number,
): { start: number; size: number } {
  switch (from) {
    case 'page':
      return { start: 0, size: width }
    case 'leftMargin':
      return { start: 0, size: page.margins.left }
    case 'rightMargin':
      return { start: width - page.margins.right, size: page.margins.right }
    case 'insideMargin':
      return { start: 0, size: page.margins.left }
    case 'outsideMargin':
      return { start: width - page.margins.right, size: page.margins.right }
    // `margin`, `column` and `character` coincide on a single-column page.
    default:
      return { start: columnLeft, size: columnWidth }
  }
}

export function floatsOf(attrs: Record<string, unknown> | null | undefined): FloatingObject[] {
  const raw = attrs?.['floats']
  return Array.isArray(raw) ? (raw as FloatingObject[]) : []
}

export interface AnchoredFloat {
  readonly object: FloatingObject
  readonly anchorTopMm: number
}

/**
 * They repeat on every sheet, like the band. A band's "paragraph" starts at the distance `w:pgMar`
 * declares: from the top in the header, from the bottom in the footer.
 */
export function bandFloatsOf(page: PageSetup, pageNumber: number): AnchoredFloat[] {
  const height = pageDimensionsMm(page).height
  const header = bandForPage(page, pageNumber, 'header')
  const footer = bandForPage(page, pageNumber, 'footer')
  const label = pageLabel(page, pageNumber)

  return [
    ...(header?.floats ?? []).map((object) => ({
      object: numbered(object, label),
      anchorTopMm: page.headerDistanceMm,
    })),
    ...(footer?.floats ?? []).map((object) => ({
      object: numbered(object, label),
      anchorTopMm: height - page.footerDistanceMm,
    })),
  ]
}

/**
 * The reader delivers the box's `PAGE` field as `{n}`; without the replacement the sheet would show
 * the braces.
 */
function numbered(object: FloatingObject, pageNumber: string): FloatingObject {
  if (object.kind !== 'text' || object.content === undefined) return object

  const content = object.content.map((node) => replaceMarkers(node, pageNumber))

  // A box with numbering stops being editable: writing this sheet's number back to the file would
  // replace the `PAGE` field with a fixed number.
  const marked = JSON.stringify(content) !== JSON.stringify(object.content)

  return { ...object, content, ...(marked ? { bid: undefined } : {}) }
}

function replaceMarkers(node: DocumentNode, pageNumber: string): DocumentNode {
  return {
    ...node,
    ...(typeof node.text === 'string' ? { text: node.text.replaceAll('{n}', pageNumber) } : {}),
    ...(Array.isArray(node.content)
      ? { content: node.content.map((child) => replaceMarkers(child, pageNumber)) }
      : {}),
  }
}
