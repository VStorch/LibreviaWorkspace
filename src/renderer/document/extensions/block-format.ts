import { Extension, type Attributes } from '@tiptap/core'

/**
 * In OOXML background, spacing and line height belong to the **paragraph**, not to a text style as
 * in Tiptap: in the corpus, `Heading1` is a red bar with white text. A body paragraph only carries
 * direct formatting, and inline beats the style rule; in lists, cells and old drafts the values
 * arrive already resolved, because no style rule reaches there.
 */

export interface BlockFormatOptions {
  types: string[]
}

// Zero means "no space before", an instruction from the document.
const pointsToCss = (value: unknown): string | null => {
  if (value === null || value === undefined) return null
  const points = Number(value)
  return Number.isFinite(points) && points >= 0 ? `${points}pt` : null
}

/**
 * Zero counts: it is the paragraph undoing the style's indent. `Number(null)` is zero, hence the
 * check first.
 */
const declaredMeasure = (value: unknown): number | null => {
  if (value === null || value === undefined || value === '') return null
  const measure = Number(value)
  return Number.isFinite(measure) ? measure : null
}

const BLOCK_ATTRIBUTES: Attributes = {
  background: {
    default: null,
    parseHTML: (element) => element.style.backgroundColor || null,
    renderHTML: (attributes) => {
      const color = attributes['background']
      return typeof color === 'string' && color.length > 0 ? { style: `background-color: ${color}` } : {}
    },
  },

  /** `w:lvlText`, often in the private use area, as Word writes Symbol and Wingdings. */
  marker: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-marker'),
    renderHTML: (attributes) => {
      const value = attributes['marker']
      if (typeof value !== 'string' || value.length === 0) return {}
      // A variable, not `list-style-type`: the item's `::before` is the only way to control the
      // distance to the text. Single quotes, because the value goes into a CSS string.
      return {
        'data-marker': value,
        style: `--marca: '${value.replace(/['\\]/g, '\\$&')}'`,
      }
    },
  },

  /** The level's `w:ind/@left`. */
  indentMm: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-indent-mm'),
    renderHTML: (attributes) => {
      const value = declaredMeasure(attributes['indentMm'])
      // Also a variable: an anchored image positions itself by the column and subtracts the indent
      // through it.
      return value !== null && value >= 0
        ? {
            'data-indent-mm': String(value),
            style: `padding-left: ${value}mm; --recuo: ${value}mm`,
          }
        : {}
    },
  },

  /** `w:ind/@right`: narrows the column and changes where the line breaks. */
  indentRightMm: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-indent-right-mm'),
    renderHTML: (attributes) => {
      const value = declaredMeasure(attributes['indentRightMm'])
      return value !== null && value >= 0
        ? {
            'data-indent-right-mm': String(value),
            style: `padding-right: ${value}mm; --recuo-direita: ${value}mm`,
          }
        : {}
    },
  },

  /**
   * Positive is `w:firstLine` and negative is `w:hanging`, the same measure with the sign flipped.
   */
  firstLineMm: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-first-line-mm'),
    renderHTML: (attributes) => {
      const value = declaredMeasure(attributes['firstLineMm'])
      return value !== null ? { 'data-first-line-mm': String(value), style: `text-indent: ${value}mm` } : {}
    },
  },

  /** The level's `w:ind/@hanging`: the distance from the marker to the text. */
  hangingMm: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-hanging-mm'),
    renderHTML: (attributes) => {
      const value = Number(attributes['hangingMm'])
      return Number.isFinite(value) && value > 0
        ? { 'data-hanging-mm': String(value), style: `--pendente: ${value}mm` }
        : {}
    },
  },

  /** The empty paragraph holding `w:sectPr` **is** the mark, and LibreOffice gives it no height. */
  sectionMark: {
    default: null,
    parseHTML: (element) => element.hasAttribute('data-section-mark') || null,
    renderHTML: (attributes) => (attributes['sectionMark'] === true ? { 'data-section-mark': '' } : {}),
  },

  /**
   * The id in `sections` of the section ending here, like `w:sectPr` in OOXML. Not carried over on
   * Enter: a repeated mark would make two sections with the same id.
   */
  sectionBreak: {
    default: null,
    keepOnSplit: false,
    parseHTML: (element) => element.getAttribute('data-section-break'),
    renderHTML: (attributes) =>
      typeof attributes['sectionBreak'] === 'string'
        ? { 'data-section-break': attributes['sectionBreak'] }
        : {},
  },

  spaceBefore: {
    default: null,
    parseHTML: (element) => element.style.marginTop || null,
    renderHTML: (attributes) => {
      const value = pointsToCss(attributes['spaceBefore'])
      return value === null ? {} : { style: `margin-top: ${value}` }
    },
  },

  spaceAfter: {
    default: null,
    parseHTML: (element) => element.style.marginBottom || null,
    renderHTML: (attributes) => {
      const value = pointsToCss(attributes['spaceAfter'])
      return value === null ? {} : { style: `margin-bottom: ${value}` }
    },
  },

  /** Line height comes from the element's font, not from the text inside it. */
  fontFamily: {
    default: null,
    parseHTML: (element) => element.style.fontFamily || null,
    renderHTML: (attributes) => {
      const family = attributes['fontFamily']
      return typeof family === 'string' && family.length > 0 ? { style: `font-family: ${family}` } : {}
    },
  },

  fontSize: {
    default: null,
    parseHTML: (element) => element.style.fontSize || null,
    renderHTML: (attributes) => {
      const size = attributes['fontSize']
      return typeof size === 'string' && size.length > 0 ? { style: `font-size: ${size}` } : {}
    },
  },

  /**
   * Word's single is the height the font asks for, which in CSS is `normal`: no factor imitates it.
   */
  lineHeight: {
    default: null,
    parseHTML: (element) => element.style.lineHeight || null,
    renderHTML: (attributes) => {
      const value = attributes['lineHeight']
      if (typeof value === 'number') {
        return Number.isFinite(value) && value > 0 ? { style: `line-height: ${value}` } : {}
      }
      return typeof value === 'string' && value.length > 0 ? { style: `line-height: ${value}` } : {}
    },
  },

  /**
   * A `w:br w:type="page"` stored **inside** the paragraph: as a node it would sit at line
   * position, invalid in the schema, and misalign indexes between screen and paper.
   */
  breakAfter: {
    default: null,
    parseHTML: (element) => (element.hasAttribute('data-break-after') ? true : null),
    renderHTML: (attributes) => (attributes['breakAfter'] === true ? { 'data-break-after': '' } : {}),
  },

  /** `w:br w:type="column"`, for the same reason as `breakAfter`; not carried over on Enter. */
  columnBreakAfter: {
    default: null,
    keepOnSplit: false,
    parseHTML: (element) => (element.hasAttribute('data-column-break') ? true : null),
    renderHTML: (attributes) => (attributes['columnBreakAfter'] === true ? { 'data-column-break': '' } : {}),
  },

  /** `w:keepNext`: does not change the look; the page end mark and the export use it. */
  keepNext: {
    default: null,
    parseHTML: (element) => element.hasAttribute('data-keep-next') || null,
    renderHTML: (attributes) => (attributes['keepNext'] === true ? { 'data-keep-next': '' } : {}),
  },

  /** `w:keepLines`. `false` is the paragraph undoing what the style turns on. */
  keepLines: {
    default: null,
    parseHTML: (element) => element.hasAttribute('data-keep-lines') || null,
    renderHTML: (attributes) => (attributes['keepLines'] === true ? { 'data-keep-lines': '' } : {}),
  },

  /** `w:widowControl`: absent means on, as in Word. */
  widowControl: {
    default: null,
    parseHTML: (element) => (element.hasAttribute('data-widows-allowed') ? false : null),
    renderHTML: (attributes) => (attributes['widowControl'] === false ? { 'data-widows-allowed': '' } : {}),
  },

  /**
   * Does not change the screen: the edited paragraph still points to the original style when
   * saving.
   */
  styleId: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-style-id'),
    renderHTML: (attributes) => {
      const id = attributes['styleId']
      return typeof id === 'string' && id.length > 0 ? { 'data-style-id': id } : {}
    },
  },
}

const LIST_ATTRIBUTES: Attributes = {
  /**
   * `w:numId`: without it the writer would write `w:numId w:val="0"`, which in OOXML means "no
   * numbering".
   */
  numId: {
    default: null,
    parseHTML: (element) => {
      const value = Number(element.getAttribute('data-num-id'))
      return Number.isInteger(value) && value > 0 ? value : null
    },
    renderHTML: (attributes) => {
      const value = Number(attributes['numId'])
      return Number.isInteger(value) && value > 0 ? { 'data-num-id': String(value) } : {}
    },
  },
}

export const BlockFormat = Extension.create<BlockFormatOptions>({
  name: 'blockFormat',

  addOptions() {
    // The list is not a block in the file, but it is in the editor tree, and without declared
    // spacing it would get the editor's.
    return { types: ['paragraph', 'heading', 'bulletList', 'orderedList'] }
  },

  addGlobalAttributes() {
    return [
      { types: this.options.types, attributes: BLOCK_ATTRIBUTES },
      // Only on lists, the only ones with numbering.
      { types: ['bulletList', 'orderedList'], attributes: LIST_ATTRIBUTES },
    ]
  },
})
