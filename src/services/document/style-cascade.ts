/**
 * The same order as `StyleResolver.cs`: document defaults, then the `basedOn` chain from the
 * farthest ancestor to the nearest. Two resolvers that disagree change the document's look on
 * reopen. What the sidecar does not resolve stays out here too: `w:rStyle`, the "invert the
 * inherited" of toggle properties, table styles and theme fonts.
 */

import { cssLineHeightOf } from './line-metrics.js'
import {
  LEGACY_STYLES,
  StyleType,
  type LineSpacing,
  type StyleCharacterFormat,
  type StyleDefinition,
  type StyleParagraphFormat,
  type StyleSheet,
} from './styles.js'

/** A style with inheritance applied. An absent field means "nobody in the chain said". */
export interface ResolvedStyle {
  readonly paragraph: StyleParagraphFormat
  readonly character: StyleCharacterFormat
}

/** `StyleResolver.MaxChainDepth`: a longer chain is a broken file. */
const MAX_CHAIN_DEPTH = 16

const cache = new WeakMap<StyleSheet, Map<string, ResolvedStyle>>()

/**
 * `null` is the default style. An id the document does not define resolves only the defaults, as in
 * Word.
 */
export function resolveStyle(sheet: StyleSheet, styleId: string | null): ResolvedStyle {
  let resolved = cache.get(sheet)
  if (resolved === undefined) {
    resolved = new Map()
    cache.set(sheet, resolved)
  }

  // Null must not share a key with any id, not even the empty one, which is how only the defaults
  // are requested (see `style-css.ts`).
  const key = styleId ?? '\u0000'
  const cached = resolved.get(key)
  if (cached !== undefined) return cached

  let paragraph: StyleParagraphFormat = { ...sheet.defaults.paragraph }
  let character: StyleCharacterFormat = { ...sheet.defaults.character }
  for (const style of chainOf(sheet, styleId ?? sheet.defaults.paragraphStyleId)) {
    paragraph = overlay(paragraph, style.paragraph)
    character = overlay(character, style.character)
  }

  const result = { paragraph, character }
  resolved.set(key, result)
  return result
}

function chainOf(sheet: StyleSheet, styleId: string | null): readonly StyleDefinition[] {
  const chain: StyleDefinition[] = []
  const seen = new Set<string>()
  let current = styleId ?? undefined

  while (current !== undefined && !seen.has(current) && chain.length < MAX_CHAIN_DEPTH) {
    const style = sheet.styles[current]
    if (style === undefined) break
    seen.add(current)
    chain.push(style)
    current = style.basedOn
  }

  return chain.reverse()
}

/**
 * Field by field: a style's silence lets the inherited value through. Since before, after and line
 * spacing are separate fields, a `w:spacing` that only redeclares the space does not erase the line
 * spacing.
 */
function overlay<T extends object>(base: T, top: T | undefined): T {
  if (top === undefined) return base
  const merged: Record<string, unknown> = { ...(base as Record<string, unknown>) }
  for (const [key, value] of Object.entries(top)) {
    if (value !== undefined) merged[key] = value
  }
  return merged as T
}

/** Without the style in the document, what the writer will add when saving (`BuiltinStyles.cs`). */
export function headingStyleOf(sheet: StyleSheet, level: number): ResolvedStyle {
  const name = `heading ${level}`
  const own = Object.values(sheet.styles).find(
    (style: StyleDefinition) => style.type === StyleType.Paragraph && style.name.toLowerCase() === name,
  )
  if (own !== undefined) return resolveStyle(sheet, own.id)

  const fallback = LEGACY_STYLES.styles[`Heading${level}`]
  const base = resolveStyle(sheet, null)
  return {
    paragraph: { ...base.paragraph, ...fallback?.paragraph },
    character: { ...base.character, ...fallback?.character },
  }
}

export interface StyledBlock {
  readonly type: { readonly name: string } | string
  readonly attrs?: Readonly<Record<string, unknown>> | null | undefined
}

/**
 * The block's style with direct formatting on top, in node units. Whoever decides by attribute
 * needs the visible value: a heading whose style says "keep with next" carries no `keepNext` on the
 * node. Without a stylesheet, the raw attributes.
 */
export function effectiveAttrs(block: StyledBlock, sheet: StyleSheet | null): Record<string, unknown> {
  const attrs: Record<string, unknown> = { ...(block.attrs ?? {}) }
  const style = blockStyleOfNode(block, sheet)
  if (style === null) return attrs

  const inherited = styleAttrsOf(style)
  for (const [name, value] of Object.entries(inherited)) {
    if (attrs[name] === null || attrs[name] === undefined) attrs[name] = value
  }
  return attrs
}

/**
 * The CSS rule's criterion: the declared id; a heading without an id, by the name `heading N`; the
 * rest, the default.
 */
export function blockStyleOfNode(block: StyledBlock, sheet: StyleSheet | null): ResolvedStyle | null {
  const type = typeof block.type === 'string' ? block.type : block.type.name
  if (sheet === null || (type !== 'paragraph' && type !== 'heading')) return null

  const attrs = block.attrs ?? {}
  const styleId = typeof attrs['styleId'] === 'string' && attrs['styleId'] !== '' ? attrs['styleId'] : null
  const level = Number(attrs['level'])
  return styleId === null && type === 'heading' && Number.isInteger(level) && level >= 1
    ? headingStyleOf(sheet, level)
    : resolveStyle(sheet, styleId)
}

/**
 * Without the document defaults: on top, they would erase what the paragraph style gave the run.
 */
export function resolveCharacterStyle(sheet: StyleSheet, styleId: string): StyleCharacterFormat {
  let character: StyleCharacterFormat = {}
  for (const style of chainOf(sheet, styleId)) character = overlay(character, style.character)
  return character
}

/** An absent field means "nobody said". */
export function styleAttrsOf({ paragraph, character }: ResolvedStyle): Record<string, unknown> {
  const attrs: Record<string, unknown> = {
    textAlign: paragraph.textAlign,
    indentMm: paragraph.indentMm,
    indentRightMm: paragraph.indentRightMm,
    firstLineMm: paragraph.firstLineMm,
    // Zero when the chain is silent, like the CSS rule: it is what is visible.
    spaceBefore: paragraph.spaceBefore ?? 0,
    spaceAfter: paragraph.spaceAfter ?? 0,
    lineHeight: lineHeightAttrOf(paragraph.lineSpacing, character.fontFamily ?? null),
    background: paragraph.background,
    keepNext: paragraph.keepNext,
    keepLines: paragraph.keepLines,
    widowControl: paragraph.widowControl,
    fontFamily: character.fontFamily,
    fontSize: character.fontSize,
  }
  for (const key of Object.keys(attrs)) if (attrs[key] === undefined) delete attrs[key]
  return attrs
}

/**
 * The CSS number already multiplied by the font's natural height; `exact` and `atLeast` in points.
 */
export function lineHeightAttrOf(spacing: LineSpacing | undefined, family: string | null): string {
  if (spacing === undefined) return cssLineHeightOf(1, family)
  if (spacing.kind === 'multiple') return cssLineHeightOf(usableFactor(spacing.factor), family)
  return `${spacing.pt}pt`
}

/**
 * Outside `(0.5; 4)` it is file garbage and single applies, the same cut as
 * `BodyReader.LineHeightOf`.
 */
export function usableFactor(factor: number): number {
  return factor > 0.5 && factor < 4 ? factor : 1
}
