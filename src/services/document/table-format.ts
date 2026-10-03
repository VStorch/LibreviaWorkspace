/**
 * Borda e sombreamento viajam como **texto canônico** no nó da célula: a
 * impressão digital é o JSON do nó (ver `Nodes.cs`), e dois objetos que descrevem
 * a mesma célula de jeitos diferentes fariam toda tabela ser regenerada.
 *
 * Só `w:tcBorders` e `w:shd/@fill`. Estilo de borda que o CSS não desenha vira
 * `single` na leitura, e só a célula que a pessoa formatou é reescrita (ver
 * `TableLook.cs`).
 */

/** Os estilos que o OOXML e o CSS desenham do mesmo jeito. */
export const CellBorderStyle = {
  /** Borda apagada de propósito — `w:val="nil"`, e não a ausência de borda. */
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
  /** Espessura em pontos. O OOXML mede em oitavos de ponto (`w:sz`). */
  readonly widthPt: number
  /** `#rrggbb` minúsculo. */
  readonly color: string
}

/** `null` num lado é "o documento não diz nada sobre este lado". */
export type CellBorders = { readonly [Side in CellBorderSide]: CellBorder | null }

export const NO_CELL_BORDERS: CellBorders = { top: null, right: null, bottom: null, left: null }

/** A espessura que o Word oferece, e o teto que o `w:sz` aceita (255 oitavos). */
export const MIN_BORDER_PT = 0.25
export const MAX_BORDER_PT = 31

const HEX = /^#[0-9a-f]{6}$/

/** Cor válida é a que o `w:color` aceita: seis dígitos hexadecimais. */
export function isCellColor(value: string): boolean {
  return HEX.test(value.toLowerCase())
}

function isBorderStyle(value: string): value is CellBorderStyle {
  return (Object.values(CellBorderStyle) as string[]).includes(value)
}

/** Ponto decimal e sem zero à direita, como os dois lados escrevem: senão, duas impressões digitais. */
function formatPt(value: number): string {
  return Number(value.toFixed(2)).toString()
}

/** O texto canônico do atributo, ou `null` quando não há lado nenhum a declarar. */
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

/** A mesma borda nos lados escolhidos, preservando os outros. */
export function withBorderOnSides(
  borders: CellBorders,
  sides: readonly CellBorderSide[],
  border: CellBorder | null,
): CellBorders {
  const next: Record<CellBorderSide, CellBorder | null> = { ...borders }
  for (const side of sides) next[side] = border
  return next
}

/** `none` vira `0`: `w:val="nil"` apaga a borda da tabela, que apareceria por baixo em `border-collapse`. */
export function cellBordersToCss(borders: CellBorders): string {
  return CELL_BORDER_SIDES.flatMap((side) => {
    const border = borders[side]
    if (border === null) return []
    if (border.style === CellBorderStyle.None) return [`border-${side}:0`]
    return [`border-${side}:${formatPt(border.widthPt)}pt ${border.style} ${border.color}`]
  }).join(';')
}

export interface TableDraft {
  /** `null` é a coluna nunca medida, de tabela recém-inserida: aplicar não mexe na largura. */
  readonly columnWidthMm: number | null
  readonly borderStyle: CellBorderStyle
  readonly borderWidthPt: number
  readonly borderColor: string
  readonly sides: { readonly [Side in CellBorderSide]: boolean }
  /** Desligado é "sem sombreamento", que no arquivo é a ausência de `w:shd`. */
  readonly shaded: boolean
  readonly shadingColor: string
  /** `w:tblHeader`: a linha se repete no alto de cada página. */
  readonly headerRow: boolean
}

/** Faixa de largura de coluna: menos de 5 mm não caberia um caractere. */
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

/** Borda e sombreamento de uma célula, como viajam no nó. */
export interface CellLook {
  readonly borders: string | null
  readonly shading: string | null
}

/**
 * Só os campos que a pessoa alterou, contados contra o rascunho **de abertura**.
 * O rascunho resume a célula numa borda só; aplicá-lo inteiro reescreveria os
 * quatro lados com o resumo, e a célula com `top: nil` e `bottom: single`
 * perderia a borda de baixo.
 */
export function cellLookPatch(cell: CellLook, before: TableDraft, after: TableDraft): CellLook {
  const styleChanged = after.borderStyle !== before.borderStyle
  const widthChanged = after.borderWidthPt !== before.borderWidthPt
  const colorChanged = after.borderColor.toLowerCase() !== before.borderColor.toLowerCase()
  const sidesChanged = CELL_BORDER_SIDES.some((side) => after.sides[side] !== before.sides[side])

  let borders = cell.borders
  if (styleChanged || widthChanged || colorChanged || sidesChanged) {
    const current = cellBordersFromAttr(cell.borders)
    const next: Record<CellBorderSide, CellBorder | null> = { ...current }

    for (const side of CELL_BORDER_SIDES) {
      const was = before.sides[side]
      const is = after.sides[side]
      const existing = current[side]

      if (was && !is) {
        next[side] = null
      } else if (!was && is) {
        next[side] = {
          style: after.borderStyle,
          widthPt: after.borderWidthPt,
          color: after.borderColor.toLowerCase(),
        }
      } else if (existing !== null) {
        next[side] = {
          style: styleChanged ? after.borderStyle : existing.style,
          widthPt: widthChanged ? after.borderWidthPt : existing.widthPt,
          color: colorChanged ? after.borderColor.toLowerCase() : existing.color,
        }
      }
    }

    borders = cellBordersToAttr(next)
  }

  const shadingChanged =
    after.shaded !== before.shaded ||
    (after.shaded && after.shadingColor.toLowerCase() !== before.shadingColor.toLowerCase())
  const shading = shadingChanged ? (after.shaded ? after.shadingColor.toLowerCase() : null) : cell.shading

  return { borders, shading }
}

/**
 * Como a tela desenha (`table-layout: fixed`): o que sobra é dividido entre as
 * colunas sem medida. Aplicar uma largura só deixaria as outras em zero, e o
 * gravador descarta grade parcial.
 */
export function resolvedColumnWidths(declared: readonly (number | null)[], totalPx: number): number[] {
  const known = declared.filter((width): width is number => width !== null && width > 0)
  const unsized = declared.length - known.length
  if (unsized === 0) return declared.map((width) => Math.max(1, Math.round(width ?? 1)))

  const remaining = totalPx - known.reduce((sum, width) => sum + width, 0)
  const share = Math.max(1, Math.round(remaining / unsized))

  return declared.map((width) => (width !== null && width > 0 ? Math.round(width) : share))
}

/** Cada célula é um parágrafo medido pela paginação: mil por mil travaria o editor. */
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
