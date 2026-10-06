/**
 * One rule per paragraph style, with **every** field explicit (zero when the chain is silent), so
 * nothing inherits the `p` margin or the `h1` bold. Direct formatting goes inline and beats any
 * selector. Only direct children of `.page__content`: lists and cells have their own rules in
 * `content-styles.ts`.
 *
 * `w:contextualSpacing` is left out: pagination checked against the corpus was measured without it,
 * and it goes back to the file intact.
 */

import { cssLineHeightOf, explicitCssLineHeightOf } from './line-metrics.js'
import {
  headingStyleOf,
  resolveCharacterStyle,
  resolveStyle,
  usableFactor,
  type ResolvedStyle,
} from './style-cascade.js'
import { StyleType, type StyleCharacterFormat, type StyleSheet } from './styles.js'

const HEADING_LEVELS = 6

const BLACK = '#000000'

export const DECLARED_BLACK_VARIABLE = '--declared-black'

export function styleSheetCss(sheet: StyleSheet): string {
  const base = resolveStyle(sheet, null)
  const rules = [
    `.page__content { ${declarations(characterCss(base)).join(' ')} }`,
    // The space between non-paragraph blocks is the default style's space before.
    `.page__content > * + * { margin-top: ${points(base.paragraph.spaceBefore)}; }`,
    rule('.page__content > p:not([data-style-id])', base),
  ]

  for (let level = 1; level <= HEADING_LEVELS; level++) {
    const style = headingStyleOf(sheet, level)
    rules.push(rule(`.page__content > h${level}:not([data-style-id])`, style))
  }

  // An id the document does not define: Word draws it with the defaults only. Before the per-id
  // rules, which win by coming later.
  rules.push(rule('.page__content > [data-style-id]', resolveStyle(sheet, '')))

  // Character styles, at any depth, with only what the chain declares.
  for (const style of Object.values(sheet.styles)) {
    if (style.type !== StyleType.Character) continue
    const css = declaredCharacterCss(resolveCharacterStyle(sheet, style.id))
    if (css.length === 0) continue
    rules.push(
      `.page__content [data-char-style="${attributeText(style.id)}"] { ${declarations(css).join(' ')} }`,
    )
  }

  for (const style of Object.values(sheet.styles)) {
    if (style.type !== StyleType.Paragraph) continue
    rules.push(
      rule(`.page__content > [data-style-id="${attributeText(style.id)}"]`, resolveStyle(sheet, style.id)),
    )
  }

  // The table of contents is a block in the editor and paragraphs with `toc 1`, `toc 2`… in the
  // file. Repeated at the end, in the same order, for the same precedence.
  const inContents = rules
    .filter((text) => text.startsWith('.page__content > ') && !text.startsWith('.page__content > * + *'))
    .map((text) => text.replace('.page__content > ', '.page__content > [data-toc] > '))

  return [...rules, ...inContents].join('\n')
}

function rule(selector: string, style: ResolvedStyle): string {
  return `${selector} { ${declarations([...paragraphCss(style), ...characterCss(style)]).join(' ')} }`
}

function paragraphCss({ paragraph }: ResolvedStyle): Array<[string, string]> {
  const indent = millimeters(paragraph.indentMm)
  const indentRight = millimeters(paragraph.indentRightMm)
  return [
    ['margin-top', points(paragraph.spaceBefore)],
    ['margin-bottom', points(paragraph.spaceAfter)],
    ['text-align', paragraph.textAlign ?? 'start'],
    ['padding-left', indent],
    ['--recuo', indent],
    ['padding-right', indentRight],
    ['--recuo-direita', indentRight],
    ['text-indent', millimeters(paragraph.firstLineMm)],
    ['background-color', paragraph.background ?? 'transparent'],
  ]
}

function characterCss({ paragraph, character }: ResolvedStyle): Array<[string, string]> {
  const family = character.fontFamily ?? null
  const decorations = [
    character.underline === true ? 'underline' : '',
    character.strike === true ? 'line-through' : '',
  ]
    .filter((line) => line.length > 0)
    .join(' ')

  const css: Array<[string, string]> = [
    ['font-size', character.fontSize ?? '12pt'],
    ['font-weight', character.bold === true ? '700' : '400'],
    ['font-style', character.italic === true ? 'italic' : 'normal'],
    ['text-decoration', decorations.length > 0 ? decorations : 'none'],
    ['text-transform', character.allCaps === true ? 'uppercase' : 'none'],
    ['font-variant', character.smallCaps === true ? 'small-caps' : 'normal'],
    ['line-height', lineHeightOf(paragraph, family)],
  ]
  if (family !== null) css.unshift(['font-family', fontStackOf(family)])
  // Without a declared color Word's "automatic" applies, which is not a color to write.
  if (character.color !== undefined) css.push(['color', screenColorOf(character.color)])
  return css
}

/** What the character style leaves unsaid belongs to the paragraph. */
function declaredCharacterCss(character: StyleCharacterFormat): Array<[string, string]> {
  const declared: Array<[string, string | undefined]> = [
    ['font-family', mapDefined(character.fontFamily, fontStackOf)],
    ['font-size', character.fontSize],
    ['font-weight', mapDefined(character.bold, (bold) => (bold ? '700' : '400'))],
    ['font-style', mapDefined(character.italic, (italic) => (italic ? 'italic' : 'normal'))],
    ['text-decoration', declaredDecoration(character)],
    ['text-transform', mapDefined(character.allCaps, (caps) => (caps ? 'uppercase' : 'none'))],
    ['font-variant', mapDefined(character.smallCaps, (small) => (small ? 'small-caps' : 'normal'))],
    ['color', mapDefined(character.color, screenColorOf)],
    ['background-color', character.highlight],
  ]
  return declared.filter((entry): entry is [string, string] => entry[1] !== undefined)
}

function screenColorOf(color: string): string {
  return color.toLowerCase() === BLACK ? `var(${DECLARED_BLACK_VARIABLE}, ${BLACK})` : color
}

function mapDefined<T>(value: T | undefined, map: (value: T) => string): string | undefined {
  return value === undefined ? undefined : map(value)
}

function declaredDecoration(character: StyleCharacterFormat): string | undefined {
  const lines = [
    character.underline === true ? 'underline' : '',
    character.strike === true ? 'line-through' : '',
  ].filter((line) => line !== '')
  return lines.length === 0 ? undefined : lines.join(' ')
}

function lineHeightOf(paragraph: ResolvedStyle['paragraph'], family: string | null): string {
  const spacing = paragraph.lineSpacing
  if (spacing === undefined) return cssLineHeightOf(1, family)
  switch (spacing.kind) {
    case 'multiple':
      return cssLineHeightOf(usableFactor(spacing.factor), family)
    case 'exact':
      return `${spacing.pt}pt`
    case 'atLeast':
      return `max(${spacing.pt}pt, ${explicitCssLineHeightOf(1, family)}em)`
  }
}

/** A bare name from the builtin tables gets quotes and a generic family. */
function fontStackOf(family: string): string {
  if (family.includes(',')) return family
  const generic = /serif|times|roman|cambria|georgia|garamond/i.test(family) ? 'serif' : 'sans-serif'
  return `'${family.replace(/'/g, '')}', ${generic}`
}

function declarations(pairs: ReadonlyArray<readonly [string, string]>): string[] {
  return pairs.map(([property, value]) => `${property}: ${value};`)
}

function points(value: number | undefined): string {
  return `${value ?? 0}pt`
}

function millimeters(value: number | undefined): string {
  return `${value ?? 0}mm`
}

/** The id goes in quotes in an attribute selector: only quotes and backslash need escaping. */
function attributeText(id: string): string {
  return id.replace(/["\\]/g, '\\$&')
}
