/**
 * Borders and shading travel as **canonical text** on the cell node: the fingerprint is the node's
 * JSON (see `Nodes.cs`), and two objects describing the same cell differently would regenerate
 * every table.
 *
 * Only `w:tcBorders` and `w:shd/@fill`. A border style CSS cannot draw becomes `single` on read,
 * and only the cell the user formatted is rewritten (see `TableLook.cs`).
 */

/** Styles OOXML and CSS draw the same way. */
export const CellBorderStyle = {
  /** A border erased on purpose: `w:val="nil"`, not the absence of a border. */
  None: 'none',
  Single: 'single',
  Double: 'double',
  Dashed: 'dashed',
  Dotted: 'dotted',
} as const
export type CellBorderStyle = (typeof CellBorderStyle)[keyof typeof CellBorderStyle]

export const CELL_BORDER_SIDES = ['top', 'right', 'bottom', 'left'] as const
export type CellBorderSide = (typeof CELL_BORDER_SIDES)[number]

export interface CellBorder {
  readonly style: CellBorderStyle
  /** In points. OOXML measures eighths of a point (`w:sz`). */
  readonly widthPt: number
  /** Lowercase `#rrggbb`. */
  readonly color: string
}

/** `null` on a side means "the document says nothing about this side". */
export type CellBorders = { readonly [Side in CellBorderSide]: CellBorder | null }

export const NO_CELL_BORDERS: CellBorders = { top: null, right: null, bottom: null, left: null }

/** The width Word offers, and the ceiling `w:sz` accepts (255 eighths). */
export const MIN_BORDER_PT = 0.25
export const MAX_BORDER_PT = 31

const HEX = /^#[0-9a-f]{6}$/

/** What `w:color` accepts: six hex digits. */
export function isCellColor(value: string): boolean {
  return HEX.test(value.toLowerCase())
}

function isBorderStyle(value: string): value is CellBorderStyle {
  return (Object.values(CellBorderStyle) as string[]).includes(value)
}

/** Decimal point and no trailing zero, as both sides write it: otherwise two fingerprints. */
function formatPt(value: number): string {
  return Number(value.toFixed(2)).toString()
}

/** `null` when there is no side to declare. */
export function cellBordersToAttr(borders: CellBorders): string | null {
  const parts = CELL_BORDER_SIDES.flatMap((side) => {
    const border = borders[side]
    if (border === null) return []
    return [`${side}:${border.style},${formatPt(border.widthPt)},${border.color}`]
  })

  return parts.length === 0 ? null : parts.join(';')
}

export function cellBordersFromAttr(value: unknown): CellBorders {
  if (typeof value !== 'string' || value === '') return NO_CELL_BORDERS

  const borders: Record<CellBorderSide, CellBorder | null> = { ...NO_CELL_BORDERS }

  for (const part of value.split(';')) {
    const [side, rest] = part.split(':')
    if (side === undefined || rest === undefined) continue
    if (!CELL_BORDER_SIDES.includes(side as CellBorderSide)) continue

    const [style, width, color] = rest.split(',')
    if (style === undefined || !isBorderStyle(style)) continue

    const widthPt = Number(width)
    borders[side as CellBorderSide] = {
      style,
      widthPt: Number.isFinite(widthPt) && widthPt > 0 ? widthPt : MIN_BORDER_PT,
      color: color !== undefined && isCellColor(color) ? color.toLowerCase() : '#000000',
    }
  }

  return borders
}

/** Preserving the other sides. */
export function withBorderOnSides(
  borders: CellBorders,
  sides: readonly CellBorderSide[],
  border: CellBorder | null,
): CellBorders {
  const next: Record<CellBorderSide, CellBorder | null> = { ...borders }
  for (const side of sides) next[side] = border
  return next
}

/**
 * `none` becomes `0`: `w:val="nil"` erases the table border, which would show through under
 * `border-collapse`.
 */
export function cellBordersToCss(borders: CellBorders): string {
  return CELL_BORDER_SIDES.flatMap((side) => {
    const border = borders[side]
    if (border === null) return []
    if (border.style === CellBorderStyle.None) return [`border-${side}:0`]
    return [`border-${side}:${formatPt(border.widthPt)}pt ${border.style} ${border.color}`]
  }).join(';')
}

export interface TableDraft {
  /**
   * `null` is a never-measured column of a freshly inserted table: applying leaves its width alone.
   */
  readonly columnWidthMm: number | null
  readonly borderStyle: CellBorderStyle
  readonly borderWidthPt: number
  readonly borderColor: string
  readonly sides: { readonly [Side in CellBorderSide]: boolean }
  /** Off means "no shading", which in the file is the absence of `w:shd`. */
  readonly shaded: boolean
  readonly shadingColor: string
  /** `w:tblHeader`: the row repeats at the top of each page. */
  readonly headerRow: boolean
}

/** Below 5 mm not even a character would fit. */
export const MIN_COLUMN_WIDTH_MM = 5
export const MAX_COLUMN_WIDTH_MM = 500

export const DEFAULT_TABLE_DRAFT: TableDraft = {
  columnWidthMm: null,
  borderStyle: CellBorderStyle.Single,
  borderWidthPt: 0.5,
  borderColor: '#000000',
  sides: { top: true, right: true, bottom: true, left: true },
  shaded: false,
  shadingColor: '#d9d9d9',
  headerRow: false,
}

export function isValidTableDraft(draft: TableDraft): boolean {
  if (!isCellColor(draft.borderColor) || !isCellColor(draft.shadingColor)) return false
  if (!Number.isFinite(draft.borderWidthPt)) return false
  if (draft.borderWidthPt < MIN_BORDER_PT || draft.borderWidthPt > MAX_BORDER_PT) return false

  if (draft.columnWidthMm === null) return true
  return (
    Number.isFinite(draft.columnWidthMm) &&
    draft.columnWidthMm >= MIN_COLUMN_WIDTH_MM &&
    draft.columnWidthMm <= MAX_COLUMN_WIDTH_MM
  )
}

export function tableDraftFrom(attrs: {
  readonly borders?: unknown
  readonly shading?: unknown
  readonly columnWidthMm?: number | null
  readonly headerRow?: boolean
}): TableDraft {
  const borders = cellBordersFromAttr(attrs.borders)
  const declared = CELL_BORDER_SIDES.map((side) => borders[side]).find((border) => border !== null)
  const shading = typeof attrs.shading === 'string' && isCellColor(attrs.shading) ? attrs.shading : null

  return {
    columnWidthMm: attrs.columnWidthMm ?? null,
    borderStyle: declared?.style ?? DEFAULT_TABLE_DRAFT.borderStyle,
    borderWidthPt: declared?.widthPt ?? DEFAULT_TABLE_DRAFT.borderWidthPt,
    borderColor: declared?.color ?? DEFAULT_TABLE_DRAFT.borderColor,
    sides: {
      top: borders.top !== null,
      right: borders.right !== null,
      bottom: borders.bottom !== null,
      left: borders.left !== null,
    },
    shaded: shading !== null,
    shadingColor: shading ?? DEFAULT_TABLE_DRAFT.shadingColor,
    headerRow: attrs.headerRow ?? false,
  }
}

/** As they travel on the node. */
export interface CellLook {
  readonly borders: string | null
  readonly shading: string | null
}

/**
 * Only the fields the user changed, compared against the **opening** draft. The draft sums the cell
 * up as a single border; applying it whole would rewrite all four sides with the summary, and a
 * cell with `top: nil` and `bottom: single` would lose its bottom border.
 */
export function cellLookPatch(cell: CellLook, before: TableDraft, after: TableDraft): CellLook {
  return {
    borders: patchedBorders(cell.borders, before, after),
    shading: patchedShading(cell.shading, before, after),
  }
}

interface BorderChanges {
  readonly style: boolean
  readonly width: boolean
  readonly color: boolean
}

function patchedBorders(borders: string | null, before: TableDraft, after: TableDraft): string | null {
  const changed: BorderChanges = {
    style: after.borderStyle !== before.borderStyle,
    width: after.borderWidthPt !== before.borderWidthPt,
    color: after.borderColor.toLowerCase() !== before.borderColor.toLowerCase(),
  }
  const sidesChanged = CELL_BORDER_SIDES.some((side) => after.sides[side] !== before.sides[side])
  if (!changed.style && !changed.width && !changed.color && !sidesChanged) return borders

  const current = cellBordersFromAttr(borders)
  const next: Record<CellBorderSide, CellBorder | null> = { ...current }
  for (const side of CELL_BORDER_SIDES) {
    const sides = { was: before.sides[side], is: after.sides[side] }
    next[side] = patchedSide(current[side], sides, after, changed)
  }
  return cellBordersToAttr(next)
}

/**
 * A side switched on gets the draft's whole border; one that already had a border changes only what
 * was altered.
 */
function patchedSide(
  existing: CellBorder | null,
  sides: { readonly was: boolean; readonly is: boolean },
  after: TableDraft,
  changed: BorderChanges,
): CellBorder | null {
  if (sides.was && !sides.is) return null
  const color = after.borderColor.toLowerCase()
  if (!sides.was && sides.is) return { style: after.borderStyle, widthPt: after.borderWidthPt, color }
  if (existing === null) return null
  return {
    style: changed.style ? after.borderStyle : existing.style,
    widthPt: changed.width ? after.borderWidthPt : existing.widthPt,
    color: changed.color ? color : existing.color,
  }
}

function patchedShading(shading: string | null, before: TableDraft, after: TableDraft): string | null {
  const changed =
    after.shaded !== before.shaded ||
    (after.shaded && after.shadingColor.toLowerCase() !== before.shadingColor.toLowerCase())
  if (!changed) return shading
  return after.shaded ? after.shadingColor.toLowerCase() : null
}

/**
 * As the screen draws (`table-layout: fixed`): what is left is shared among unmeasured columns.
 * Applying a single width would leave the others at zero, and the writer drops a partial grid.
 */
export function resolvedColumnWidths(declared: readonly (number | null)[], totalPx: number): number[] {
  const known = declared.filter((width): width is number => width !== null && width > 0)
  const unsized = declared.length - known.length
  if (unsized === 0) return declared.map((width) => Math.max(1, Math.round(width ?? 1)))

  const remaining = totalPx - known.reduce((sum, width) => sum + width, 0)
  const share = Math.max(1, Math.round(remaining / unsized))

  return declared.map((width) => (width !== null && width > 0 ? Math.round(width) : share))
}

/**
 * Each cell is a paragraph measured by pagination: a thousand by a thousand would freeze the
 * editor.
 */
export const MAX_TABLE_ROWS = 100
export const MAX_TABLE_COLUMNS = 40

export function isValidTableSize(rows: number, columns: number): boolean {
  return (
    Number.isInteger(rows) &&
    Number.isInteger(columns) &&
    rows >= 1 &&
    columns >= 1 &&
    rows <= MAX_TABLE_ROWS &&
    columns <= MAX_TABLE_COLUMNS
  )
}
