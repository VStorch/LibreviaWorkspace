/**
 * The `word/styles.xml` definitions, outside the nodes (see `model.ts`).
 *
 * Units are those of block attributes, except **line spacing**: here the factor is raw, because
 * with inheritance the font may come from the style itself, and only after the cascade is resolved
 * does `line-metrics.ts` multiply it by the natural height.
 */

import { Language, translate } from '@shared/i18n/index.js'

export type LineSpacing =
  /** Times the natural line height (`w:lineRule="auto"`). */
  | { readonly kind: 'multiple'; readonly factor: number }
  /** Exactly this height, clipping what does not fit (`w:lineRule="exact"`). */
  | { readonly kind: 'exact'; readonly pt: number }
  /** This height or more (`w:lineRule="atLeast"`). */
  | { readonly kind: 'atLeast'; readonly pt: number }

/**
 * Absent means "this style does not say", and the inherited value passes; zero is the style saying
 * "no space". `| undefined` because zod emits the key with `undefined`, and
 * `exactOptionalPropertyTypes` would refuse it.
 */
export interface StyleParagraphFormat {
  readonly textAlign?: string | undefined
  /** Left indent in millimetres, like the page setup ruler. */
  readonly indentMm?: number | undefined
  readonly indentRightMm?: number | undefined
  /** First line: positive goes in, negative goes out (Word's hanging). */
  readonly firstLineMm?: number | undefined
  /** In points, as Word shows them. */
  readonly spaceBefore?: number | undefined
  readonly spaceAfter?: number | undefined
  readonly lineSpacing?: LineSpacing | undefined
  readonly keepNext?: boolean | undefined
  readonly keepLines?: boolean | undefined
  /** `w:widowControl`: absent means on, as in Word. */
  readonly widowControl?: boolean | undefined
  readonly pageBreakBefore?: boolean | undefined
  readonly contextualSpacing?: boolean | undefined
  /** Outline level 0 to 8, which makes a style a heading; 9 is body text. */
  readonly outlineLevel?: number | undefined
  readonly background?: string | undefined
}

export interface StyleCharacterFormat {
  /** A CSS stack, built from the font table. */
  readonly fontFamily?: string | undefined
  /** A measure with a unit, like the block attribute: `12pt`. */
  readonly fontSize?: string | undefined
  readonly bold?: boolean | undefined
  readonly italic?: boolean | undefined
  readonly underline?: boolean | undefined
  readonly strike?: boolean | undefined
  readonly allCaps?: boolean | undefined
  readonly smallCaps?: boolean | undefined
  readonly verticalAlign?: string | undefined
  readonly color?: string | undefined
  readonly highlight?: string | undefined
}

export const StyleType = {
  Paragraph: 'paragraph',
  Character: 'character',
} as const
export type StyleType = (typeof StyleType)[keyof typeof StyleType]

/** As the file declares it, without resolved inheritance. */
export interface StyleDefinition {
  /** `w:styleId`: what the paragraph points to. It is translated (`Ttulo1` in Portuguese). */
  readonly id: string
  /** `w:name`: the internal name, which is **not** translated (`heading 1` in any language). */
  readonly name: string
  readonly type: StyleType
  /** A recommended style, which Word shows in the gallery (`w:qFormat`). */
  readonly qFormat: boolean
  /** Hidden from the list: `w:hidden` (always) or `w:semiHidden` (until used). */
  readonly hidden: boolean
  /** Created by the document's author, not built into Word (`w:customStyle`). */
  readonly custom: boolean
  readonly basedOn?: string | undefined
  readonly next?: string | undefined
  /** `w:link`. */
  readonly link?: string | undefined
  readonly uiPriority?: number | undefined
  readonly paragraph?: StyleParagraphFormat | undefined
  readonly character?: StyleCharacterFormat | undefined
}

/**
 * Without the ids, a document that calls the default style `Padro` does not say which style the
 * paragraph has.
 */
export interface StyleDefaults {
  readonly paragraph: StyleParagraphFormat
  readonly character: StyleCharacterFormat
  readonly paragraphStyleId: string | null
  readonly characterStyleId: string | null
}

export interface StyleSheet {
  readonly defaults: StyleDefaults
  readonly styles: Readonly<Record<string, StyleDefinition>>
}

/**
 * 1.5 ÷ 1.1499 (Liberation Serif's height, see `line-metrics.ts`): the file's multiple is measured
 * against the font height, not the size. Four decimals because `w:line` lives on a 240ths grid.
 */
export const BODY_LINE_FACTOR = 1.3042

const HEADING_SIZES = ['22pt', '17pt', '14pt', '12pt', '10pt', '8pt'] as const

/**
 * The heading's space before is CSS's `margin-top: 1em` (0.6em on `h5` and `h6`, which have no rule
 * of their own), in points over each one's size.
 */
const HEADING_BEFORE = [22, 17, 14, 12, 6, 4.8] as const

/**
 * The space after is the browser's bottom margin: 0.67em on `h1`, 0.83em on `h2`, 1em on `h3`,
 * 1.33em on `h4`, 1.67em and 2.33em on the last two.
 */
const HEADING_AFTER = [14.75, 14.1, 14, 15.95, 16.65, 18.75] as const

/** "Keep with next" is print's `break-after: avoid`, which only the first four have. */
const HEADING_KEEP_NEXT = 4

function heading(level: number): StyleDefinition {
  const index = level - 1
  const paragraph: StyleParagraphFormat = {
    spaceBefore: HEADING_BEFORE[index]!,
    spaceAfter: HEADING_AFTER[index]!,
    ...(level <= HEADING_KEEP_NEXT ? { keepNext: true } : {}),
    outlineLevel: index,
  }

  return {
    id: `Heading${level}`,
    // The internal English name: reader and writer recognize a heading by it in any language.
    name: `heading ${level}`,
    type: StyleType.Paragraph,
    qFormat: true,
    hidden: false,
    custom: false,
    basedOn: 'Normal',
    next: 'Normal',
    uiPriority: 9,
    paragraph,
    character: { fontSize: HEADING_SIZES[index]!, bold: true },
  }
}

/**
 * The styles of a `.sdoc` before version 3: the CSS the editor drew before styles (Times New Roman
 * 12 pt, 1.5 spacing) plus the browser default for headings. A different number makes the old
 * document reopen with other pagination. `src/main/sidecar/builtin-styles.test.ts` compares
 * `BuiltinStyles.cs` with this table.
 *
 * Without the text's `#111111`: once saved it would come back as an explicit color.
 */
export const LEGACY_STYLES: StyleSheet = {
  defaults: {
    paragraph: {},
    // In the document defaults, not in `Normal`: that is where Word looks for the font.
    character: { fontFamily: 'Times New Roman', fontSize: '12pt' },
    paragraphStyleId: 'Normal',
    characterStyleId: 'DefaultParagraphFont',
  },
  styles: {
    Normal: {
      id: 'Normal',
      name: 'Normal',
      type: StyleType.Paragraph,
      qFormat: true,
      hidden: false,
      custom: false,
      paragraph: {
        // 0.6em before (`.page__content > * + *`) and 1em after (the browser default for `p`), over
        // 12 pt.
        spaceBefore: 7.2,
        spaceAfter: 12,
        lineSpacing: { kind: 'multiple', factor: BODY_LINE_FACTOR },
      },
    },
    DefaultParagraphFont: {
      id: 'DefaultParagraphFont',
      name: 'Default Paragraph Font',
      type: StyleType.Character,
      qFormat: false,
      hidden: true,
      custom: false,
      uiPriority: 1,
    },
    Heading1: heading(1),
    Heading2: heading(2),
    Heading3: heading(3),
    Heading4: heading(4),
    Heading5: heading(5),
    Heading6: heading(6),
    ListParagraph: {
      id: 'ListParagraph',
      name: 'List Paragraph',
      type: StyleType.Paragraph,
      qFormat: true,
      hidden: false,
      custom: false,
      basedOn: 'Normal',
      uiPriority: 34,
      // Half an inch, the OOXML indent step (720 twips).
      paragraph: { indentMm: 12.7, contextualSpacing: true },
    },
    Hyperlink: {
      id: 'Hyperlink',
      name: 'Hyperlink',
      type: StyleType.Character,
      qFormat: false,
      hidden: false,
      custom: false,
      basedOn: 'DefaultParagraphFont',
      uiPriority: 99,
      character: { color: '#0563c1', underline: true },
    },
  },
}

/**
 * Word 2013–2021 line spacing, which its dialog shows as 1.08: 259 of 240ths, written as the reader
 * returns it, on the same grid as `BODY_LINE_FACTOR`.
 */
const WORD_LINE_FACTOR = 1.0792

/** Size, space before and color of each Word 2013–2021 heading. */
const WORD_HEADINGS = [
  { fontSize: '16pt', spaceBefore: 12, color: '#2f5496' },
  { fontSize: '13pt', spaceBefore: 2, color: '#2f5496' },
  { fontSize: '12pt', spaceBefore: 2, color: '#1f3763' },
  { spaceBefore: 2, color: '#2f5496', italic: true },
  { spaceBefore: 2, color: '#2f5496' },
  { spaceBefore: 2, color: '#1f3763' },
] as const

function wordHeading(level: number): StyleDefinition {
  const { spaceBefore, ...character } = WORD_HEADINGS[level - 1]!
  return {
    ...heading(level),
    // Without Word's Calibri Light: it has no free metric-compatible substitute, and a heading
    // measured with another font breaks the line elsewhere. The heading inherits Calibri, which
    // Carlito draws the same.
    paragraph: { spaceBefore, spaceAfter: 0, keepNext: true, keepLines: true, outlineLevel: level - 1 },
    character,
  }
}

/** Word 2013–2021 default: Calibri 11 pt, 1.08 spacing, 8 pt after. */
export const BUILTIN_STYLES: StyleSheet = {
  defaults: {
    paragraph: { spaceAfter: 8, lineSpacing: { kind: 'multiple', factor: WORD_LINE_FACTOR } },
    character: { fontFamily: 'Calibri', fontSize: '11pt' },
    paragraphStyleId: 'Normal',
    characterStyleId: 'DefaultParagraphFont',
  },
  styles: {
    ...LEGACY_STYLES.styles,
    Normal: {
      id: 'Normal',
      name: 'Normal',
      type: StyleType.Paragraph,
      qFormat: true,
      hidden: false,
      custom: false,
    },
    Heading1: wordHeading(1),
    Heading2: wordHeading(2),
    Heading3: wordHeading(3),
    Heading4: wordHeading(4),
    Heading5: wordHeading(5),
    Heading6: wordHeading(6),
  },
}

/**
 * Only headings, whose name we write in English because Word requires it. Any other style's name
 * belongs to the document and shows as is.
 */
const HEADING_LABELS: Readonly<Record<string, number>> = {
  'heading 1': 1,
  'heading 2': 2,
  'heading 3': 3,
  'heading 4': 4,
  'heading 5': 5,
  'heading 6': 6,
}

export function styleLabelOf(style: StyleDefinition, language: Language = Language.Portuguese): string {
  const heading = HEADING_LABELS[style.name.toLowerCase()]
  return heading === undefined
    ? style.name
    : translate(language, 'document.styles.heading', { level: heading })
}

/**
 * Without hidden ones (`w:semiHidden` is Word's machinery). In Word's order: declared priority,
 * with the name as tiebreaker so the order is stable.
 */
export function listedStyles(
  sheet: StyleSheet,
  language: Language = Language.Portuguese,
): readonly StyleDefinition[] {
  return Object.values(sheet.styles)
    .filter((style) => !style.hidden)
    .sort((left, right) => {
      const byPriority = (left.uiPriority ?? 100) - (right.uiPriority ?? 100)
      return byPriority !== 0
        ? byPriority
        : styleLabelOf(left, language).localeCompare(
            styleLabelOf(right, language),
            language === Language.Portuguese ? 'pt-BR' : 'en-US',
          )
    })
}

export interface BlockStyleQuery {
  /** `paragraph`, `heading`, `codeBlock`… */
  readonly type: string
  /** The `w:pStyle` the block brought from the file, if any. */
  readonly styleId?: string | null | undefined
  readonly level?: number | null | undefined
}

/**
 * The block's `w:pStyle` if the document defines that style; otherwise the heading by internal name
 * (`heading 3`), because in a German document the id is `berschrift3`; otherwise the default style.
 * `null` when there is no honest answer.
 */
export function blockStyleOf(sheet: StyleSheet, block: BlockStyleQuery): StyleDefinition | null {
  const declared = block.styleId ?? null
  if (declared !== null && declared in sheet.styles) return sheet.styles[declared] ?? null

  if (block.type === 'heading' && typeof block.level === 'number') {
    const name = `heading ${block.level}`
    const found = Object.values(sheet.styles).find(
      (style) => style.type === StyleType.Paragraph && style.name.toLowerCase() === name,
    )
    if (found !== undefined) return found
  }

  const fallback = sheet.defaults.paragraphStyleId
  return (fallback !== null ? sheet.styles[fallback] : undefined) ?? null
}
