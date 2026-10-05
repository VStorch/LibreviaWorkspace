/**
 * The OOXML multiple (`w:line` with `w:lineRule="auto"`) is measured against the font's **natural
 * height** (`(ascender - descender + lineGap) / unitsPerEm` from the `hhea` table), as in Word,
 * LibreOffice and `line-height: normal`. So the `lineHeight` attribute stores the CSS number,
 * already multiplied: the reader multiplies, the writer divides, and the UI converts for display.
 *
 * `line-metrics.test.ts` compares this table with `LineMetrics.cs`. Only the installer's fonts and
 * the ones they substitute: for the rest there is no honest guess.
 */

/** The editor font when the document names none. */
export const DEFAULT_NATURAL_LINE_HEIGHT = 1.1499

const LIBERATION_SERIF = DEFAULT_NATURAL_LINE_HEIGHT

/** Keys compare ignoring case and spaces. */
export const NATURAL_LINE_HEIGHTS: Readonly<Record<string, number>> = {
  arial: 1.1499,
  helvetica: 1.1499,
  'liberation sans': 1.1499,
  'times new roman': LIBERATION_SERIF,
  'liberation serif': LIBERATION_SERIF,
  'courier new': 1.1328,
  'liberation mono': 1.1328,
  calibri: 1.2207,
  carlito: 1.2207,
  cambria: 1.15,
  caladea: 1.15,
}

/**
 * The height is the first font's; the generic one only counts if it is missing. Mirrors
 * `ParagraphFormat.FirstFont`.
 */
export function firstFontOf(stack: string | null | undefined): string | null {
  if (typeof stack !== 'string') return null
  const first =
    stack
      .split(',')[0]
      ?.trim()
      .replace(/^['"]|['"]$/g, '') ?? ''
  return first.length > 0 ? first : null
}

/** Mirrors `LineMetrics.Of`: an empty stack is the editor font; an unknown font is `null`. */
export function naturalLineHeightOf(stack: string | null | undefined): number | null {
  const first = firstFontOf(stack)
  if (first === null) return DEFAULT_NATURAL_LINE_HEIGHT
  return NATURAL_LINE_HEIGHTS[first.toLowerCase()] ?? null
}

/**
 * Mirrors `BodyReader.Multiple`: with an unknown font, factor 1 becomes `normal` and the rest uses
 * the 1.15 guess, the same as the writer.
 */
export function cssLineHeightOf(factor: number, stack: string | null | undefined): string {
  const natural = naturalLineHeightOf(stack)
  if (natural !== null) return numberText(factor * natural)
  return factor === 1 ? 'normal' : numberText(factor * DEFAULT_NATURAL_LINE_HEIGHT)
}

/**
 * To override a declared line spacing: the writer does not write `normal`, and the old `w:line`
 * would stay.
 */
export function explicitCssLineHeightOf(factor: number, stack: string | null | undefined): string {
  return numberText(factor * (naturalLineHeightOf(stack) ?? DEFAULT_NATURAL_LINE_HEIGHT))
}

/**
 * To the hundredth, the precision that survives the round trip: otherwise "1.5" would come back as
 * 1.4999.
 */
export function lineFactorOf(css: number, stack: string | null | undefined): number {
  const natural = naturalLineHeightOf(stack) ?? DEFAULT_NATURAL_LINE_HEIGHT
  return Math.round((css / natural) * 100) / 100
}

/** At most four decimals, no trailing zero: the form the reader writes. */
function numberText(value: number): string {
  return String(Math.round(value * 10_000) / 10_000)
}
