/**
 * Translation between what the block stores and what the paragraph dialog shows.
 *
 * - Line spacing is a single field, as CSS writes it (`normal`, `1.8311`, `14pt`), and two
 *   questions in the dialog: the kind and the amount. The CSS number is **not** Word's factor: it
 *   is already multiplied by the font's natural height (see `line-metrics.ts`).
 * - First-line indent is a signed number (negative `text-indent` is Word's hanging); in the
 *   dialog, a choice and a positive measure.
 */

import { INDENT_STEP_MM } from '@services/units.js'
import { cssLineHeightOf, explicitCssLineHeightOf, lineFactorOf } from './line-metrics.js'

export const LineSpacingKind = {
  /** The height the font itself asks for: the file's silence. */
  Single: 'single',
  /** A multiple of the natural height: 1.15, 1.5, double. */
  Multiple: 'multiple',
  /**
   * "At least" (`w:lineRule="atLeast"`), not "exactly": `exact` clips what does not fit, and losing
   * half a line is losing content.
   */
  AtLeast: 'at-least',
} as const

export type LineSpacingKind = (typeof LineSpacingKind)[keyof typeof LineSpacingKind]

export const FirstLineKind = {
  None: 'none',
  Indent: 'indent',
  Hanging: 'hanging',
} as const

export type FirstLineKind = (typeof FirstLineKind)[keyof typeof FirstLineKind]

export const TextAlignment = {
  Left: 'left',
  Center: 'center',
  Right: 'right',
  Justify: 'justify',
} as const

export type TextAlignment = (typeof TextAlignment)[keyof typeof TextAlignment]

/** Already numbers, never `<input>` text. */
export interface ParagraphDraft {
  readonly align: TextAlignment
  /** Points, as Word shows them. */
  readonly spaceBefore: number
  readonly spaceAfter: number
  readonly lineSpacingKind: LineSpacingKind
  /** A factor when the kind is multiple; points when it is "at least". */
  readonly lineSpacingValue: number
  /** Millimetres, like the page setup ruler. */
  readonly indentLeftMm: number
  readonly indentRightMm: number
  readonly firstLineKind: FirstLineKind
  readonly firstLineMm: number
  readonly keepNext: boolean
  readonly keepLines: boolean
  /** Widow and orphan control: on unless something says otherwise, as in Word. */
  readonly widowControl: boolean
}

export const MAX_SPACING_PT = 600
export const MAX_INDENT_MM = 200
/**
 * The writer refuses factors outside `(0.5; 4)` and records a loss (see
 * `ParagraphFormat.ApplyLineHeight`); the dialog does not offer what it cannot write.
 */
export const MIN_LINE_FACTOR = 0.51
export const MAX_LINE_FACTOR = 3.99

export const DEFAULT_PARAGRAPH_DRAFT: ParagraphDraft = {
  align: TextAlignment.Left,
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacingKind: LineSpacingKind.Single,
  lineSpacingValue: 1.15,
  indentLeftMm: 0,
  indentRightMm: 0,
  firstLineKind: FirstLineKind.None,
  firstLineMm: 0,
  keepNext: false,
  keepLines: false,
  widowControl: true,
}

/**
 * A null value clears the attribute. The dialog writes only what the user changed relative to what
 * is visible: turning inherited into direct would detach the block from its style unasked.
 */
export interface ParagraphAttrs {
  readonly textAlign: TextAlignment | null
  readonly spaceBefore: number | null
  readonly spaceAfter: number | null
  readonly lineHeight: string | null
  readonly indentMm: number | null
  readonly indentRightMm: number | null
  readonly firstLineMm: number | null
  readonly keepNext: boolean | null
  readonly keepLines: boolean | null
  readonly widowControl: boolean | null
  /**
   * Indent in `Ctrl+]` steps. The writer adds measure and level; when the dialog changes the
   * indent, the level is reset so it does not count twice.
   */
  readonly indent: number
}

const clamp = (value: number, min: number, max: number): number =>
  Number.isFinite(value) ? Math.min(Math.max(value, min), max) : min

/** The precision the reader delivers. */
const round = (value: number): number => Math.round(value * 10) / 10

function numberOf(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : null
}

/** Raw attributes, because an imported block carries measures no dialog field created. */
export function paragraphDraftFrom(attrs: Record<string, unknown>): ParagraphDraft {
  const firstLine = numberOf(attrs['firstLineMm']) ?? 0
  const indentLevel = numberOf(attrs['indent']) ?? 0

  return {
    align: alignmentOf(attrs['textAlign']),
    spaceBefore: clamp(round(numberOf(attrs['spaceBefore']) ?? 0), 0, MAX_SPACING_PT),
    spaceAfter: clamp(round(numberOf(attrs['spaceAfter']) ?? 0), 0, MAX_SPACING_PT),
    ...lineSpacingOf(attrs['lineHeight'], fontStackOf(attrs)),
    // The `Ctrl+]` level enters as a measure, so the field shows the visible indent.
    indentLeftMm: clamp(
      round((numberOf(attrs['indentMm']) ?? 0) + indentLevel * INDENT_STEP_MM),
      0,
      MAX_INDENT_MM,
    ),
    indentRightMm: clamp(round(numberOf(attrs['indentRightMm']) ?? 0), 0, MAX_INDENT_MM),
    firstLineKind:
      firstLine > 0 ? FirstLineKind.Indent : firstLine < 0 ? FirstLineKind.Hanging : FirstLineKind.None,
    firstLineMm: clamp(round(Math.abs(firstLine)), 0, MAX_INDENT_MM),
    keepNext: attrs['keepNext'] === true,
    keepLines: attrs['keepLines'] === true,
    widowControl: attrs['widowControl'] !== false,
  }
}

function alignmentOf(value: unknown): TextAlignment {
  const found = Object.values(TextAlignment).find((align) => align === value)
  return found ?? TextAlignment.Left
}

function fontStackOf(attrs: Record<string, unknown>): string | null {
  const stack = attrs['fontFamily']
  return typeof stack === 'string' && stack.trim() !== '' ? stack : null
}

function lineSpacingOf(
  value: unknown,
  fontStack: string | null,
): Pick<ParagraphDraft, 'lineSpacingKind' | 'lineSpacingValue'> {
  if (typeof value !== 'string' || value === '' || value.toLowerCase() === 'normal') {
    return {
      lineSpacingKind: LineSpacingKind.Single,
      lineSpacingValue: DEFAULT_PARAGRAPH_DRAFT.lineSpacingValue,
    }
  }

  if (value.toLowerCase().endsWith('pt')) {
    const points = numberOf(value.slice(0, -2))
    return points === null || points <= 0
      ? {
          lineSpacingKind: LineSpacingKind.Single,
          lineSpacingValue: DEFAULT_PARAGRAPH_DRAFT.lineSpacingValue,
        }
      : {
          lineSpacingKind: LineSpacingKind.AtLeast,
          lineSpacingValue: clamp(round(points), 1, MAX_SPACING_PT),
        }
  }

  const css = numberOf(value)
  // The attribute number is a CSS measure, and the dialog's is Word's line: a Calibri paragraph
  // without declared spacing arrives as 1.2207 and is "Single".
  const factor = css === null || css <= 0 ? 1 : lineFactorOf(css, fontStack)

  return factor === 1
    ? {
        lineSpacingKind: LineSpacingKind.Single,
        lineSpacingValue: DEFAULT_PARAGRAPH_DRAFT.lineSpacingValue,
      }
    : {
        lineSpacingKind: LineSpacingKind.Multiple,
        lineSpacingValue: clamp(factor, MIN_LINE_FACTOR, MAX_LINE_FACTOR),
      }
}

/** Hanging becomes a negative indent, like Word's `w:hanging`. */
export function signedFirstLineMm(draft: ParagraphDraft): number {
  if (draft.firstLineKind === FirstLineKind.None) return 0
  return draft.firstLineKind === FirstLineKind.Hanging
    ? -Math.abs(draft.firstLineMm)
    : Math.abs(draft.firstLineMm)
}

/**
 * `effective` is the block's value with the style underneath (`effectiveAttrs`), which the form was
 * built from. A field group equal to the opening one returns the raw attribute untouched; a
 * different one becomes direct formatting. Zero only becomes an attribute when it undoes something
 * from the style.
 */
export function paragraphAttrsFrom(
  draft: ParagraphDraft,
  attrs: Record<string, unknown> = {},
  effective: Record<string, unknown> = attrs,
): ParagraphAttrs {
  const shown = paragraphDraftFrom(effective)
  const field = <T>(keys: readonly (keyof ParagraphDraft)[], name: string, changed: () => T): T | null =>
    keys.every((key) => draft[key] === shown[key]) ? ((attrs[name] ?? null) as T | null) : changed()
  const indentSame = draft.indentLeftMm === shown.indentLeftMm

  return {
    textAlign: field(['align'], 'textAlign', () => keptAlignment(draft.align, effective)),
    spaceBefore: field(['spaceBefore'], 'spaceBefore', () =>
      clamp(round(draft.spaceBefore), 0, MAX_SPACING_PT),
    ),
    spaceAfter: field(['spaceAfter'], 'spaceAfter', () => clamp(round(draft.spaceAfter), 0, MAX_SPACING_PT)),
    lineHeight: field(['lineSpacingKind', 'lineSpacingValue'], 'lineHeight', () =>
      lineHeightOf(draft, effective),
    ),
    indentMm: field(['indentLeftMm'], 'indentMm', () =>
      measureOf(draft.indentLeftMm, effective, 'indentMm', 0),
    ),
    indentRightMm: field(['indentRightMm'], 'indentRightMm', () =>
      measureOf(draft.indentRightMm, effective, 'indentRightMm', 0),
    ),
    firstLineMm: field(['firstLineKind', 'firstLineMm'], 'firstLineMm', () =>
      measureOf(signedFirstLineMm(draft), effective, 'firstLineMm', -MAX_INDENT_MM),
    ),
    keepNext: field(['keepNext'], 'keepNext', () => draft.keepNext),
    keepLines: field(['keepLines'], 'keepLines', () => draft.keepLines),
    widowControl: field(['widowControl'], 'widowControl', () => draft.widowControl),
    indent: indentSame ? (numberOf(attrs['indent']) ?? 0) : 0,
  }
}

/** Zero means absent, unless the style indents: then it is an explicit zero that undoes it. */
function measureOf(
  value: number,
  effective: Record<string, unknown>,
  name: string,
  min: number,
): number | null {
  if (value !== 0) return clamp(round(value), min, MAX_INDENT_MM)
  const inherited = numberOf(effective[name]) ?? 0
  return inherited === 0 ? null : 0
}

/**
 * "Left" is only written when something declares another alignment: clearing it would let the style
 * come back.
 */
function keptAlignment(align: TextAlignment, effective: Record<string, unknown>): TextAlignment | null {
  const declared = typeof effective['textAlign'] === 'string' ? effective['textAlign'] : null
  return align === TextAlignment.Left && declared === null ? null : align
}

/** The font gives the natural line height, and what was already written decides single spacing. */
function lineHeightOf(draft: ParagraphDraft, attrs: Record<string, unknown>): string {
  if (draft.lineSpacingKind === LineSpacingKind.AtLeast) {
    return `${clamp(round(draft.lineSpacingValue), 1, MAX_SPACING_PT)}pt`
  }

  const fontStack = fontStackOf(attrs)

  if (draft.lineSpacingKind === LineSpacingKind.Single) {
    const current = attrs['lineHeight']

    // The file's single spacing is its silence: replacing it with a number would rewrite the block
    // without changing anything visible.
    if (current === null || current === undefined || current === '') return 'normal'
    if (typeof current === 'string' && current.toLowerCase() === 'normal') return 'normal'

    // Already factor 1, in the reader's form: keep it.
    const css = numberOf(current)
    if (css !== null && css > 0 && lineFactorOf(css, fontStack) === 1) return String(current)

    // A measure was declared: `normal` will not do, because the writer does not write it and the
    // old `w:line` would stay.
    return explicitCssLineHeightOf(1, fontStack)
  }

  return cssLineHeightOf(clamp(draft.lineSpacingValue, MIN_LINE_FACTOR, MAX_LINE_FACTOR), fontStack)
}

/**
 * For the toolbar and the `Ctrl+1`, `Ctrl+5` and `Ctrl+2` shortcuts, which speak in lines, as Word
 * does: `''` is single.
 */
export function lineSpacingChoiceOf(attrs: Record<string, unknown>): string {
  const spacing = lineSpacingOf(attrs['lineHeight'], fontStackOf(attrs))

  if (spacing.lineSpacingKind === LineSpacingKind.Single) return ''
  if (spacing.lineSpacingKind === LineSpacingKind.AtLeast) return `${spacing.lineSpacingValue}pt`
  return String(spacing.lineSpacingValue)
}

export function lineHeightAttrFrom(choice: string, attrs: Record<string, unknown>): string {
  const spacing = choice.endsWith('pt')
    ? { lineSpacingKind: LineSpacingKind.AtLeast, lineSpacingValue: numberOf(choice.slice(0, -2)) ?? 12 }
    : choice.trim() === ''
      ? { lineSpacingKind: LineSpacingKind.Single, lineSpacingValue: 1 }
      : { lineSpacingKind: LineSpacingKind.Multiple, lineSpacingValue: numberOf(choice) ?? 1 }

  return lineHeightOf({ ...DEFAULT_PARAGRAPH_DRAFT, ...spacing }, attrs)
}

/** Only the numbers can go out of range: the rest is a closed choice. */
export function isValidParagraphDraft(draft: ParagraphDraft): boolean {
  const inRange = (value: number, min: number, max: number): boolean =>
    Number.isFinite(value) && value >= min && value <= max

  if (!inRange(draft.spaceBefore, 0, MAX_SPACING_PT)) return false
  if (!inRange(draft.spaceAfter, 0, MAX_SPACING_PT)) return false
  if (!inRange(draft.indentLeftMm, 0, MAX_INDENT_MM)) return false
  if (!inRange(draft.indentRightMm, 0, MAX_INDENT_MM)) return false
  if (!inRange(draft.firstLineMm, 0, MAX_INDENT_MM)) return false

  if (draft.lineSpacingKind === LineSpacingKind.Multiple) {
    return inRange(draft.lineSpacingValue, MIN_LINE_FACTOR, MAX_LINE_FACTOR)
  }
  if (draft.lineSpacingKind === LineSpacingKind.AtLeast) {
    return inRange(draft.lineSpacingValue, 1, MAX_SPACING_PT)
  }

  return true
}
