import { describe, expect, it } from 'vitest'
import { currentEntryIndex, outlineLevelOf, outlineOf, type OutlineBlock } from './outline.js'
import { BUILTIN_STYLES, StyleType, type StyleSheet } from './styles.js'

/**
 * Word's sheet with an author style that is a heading, and `TOC Heading` switching the level off.
 */
const SHEET: StyleSheet = {
  ...BUILTIN_STYLES,
  styles: {
    ...BUILTIN_STYLES.styles,
    Capitulo: {
      id: 'Capitulo',
      name: 'Capítulo',
      type: StyleType.Paragraph,
      qFormat: true,
      hidden: false,
      custom: true,
      paragraph: { outlineLevel: 0 },
    },
    Secao: {
      id: 'Secao',
      name: 'Seção',
      type: StyleType.Paragraph,
      qFormat: true,
      hidden: false,
      custom: true,
      basedOn: 'Capitulo',
      paragraph: { outlineLevel: 1 },
    },
    TOCHeading: {
      id: 'TOCHeading',
      name: 'TOC Heading',
      type: StyleType.Paragraph,
      qFormat: true,
      hidden: false,
      custom: false,
      basedOn: 'Capitulo',
      paragraph: { outlineLevel: 9 },
    },
  },
}

function block(type: string, attrs: Record<string, unknown>, text: string, pos: number): OutlineBlock {
  return { type, attrs, text, pos }
}

describe('outlineLevelOf', () => {
  it('a heading counts by the level it carries', () => {
    expect(outlineLevelOf({ type: 'heading', attrs: { level: 2 } }, null)).toBe(2)
  })

  it('a paragraph counts by the level the style cascade gives it', () => {
    expect(outlineLevelOf({ type: 'paragraph', attrs: { styleId: 'Capitulo' } }, SHEET)).toBe(1)
    // The style's own level beats the inherited one.
    expect(outlineLevelOf({ type: 'paragraph', attrs: { styleId: 'Secao' } }, SHEET)).toBe(2)
  })

  it('level 9 is body text, even if it inherits from a heading', () => {
    expect(outlineLevelOf({ type: 'paragraph', attrs: { styleId: 'TOCHeading' } }, SHEET)).toBeNull()
  })

  it('a plain paragraph, or one without a style sheet, is not a heading', () => {
    expect(outlineLevelOf({ type: 'paragraph', attrs: {} }, SHEET)).toBeNull()
    expect(outlineLevelOf({ type: 'paragraph', attrs: { styleId: 'Capitulo' } }, null)).toBeNull()
  })
})

describe('outlineOf', () => {
  it('lists the headings in order, without empty ones and with whitespace normalized', () => {
    const entries = outlineOf(
      [
        block('paragraph', {}, 'Capa', 0),
        block('heading', { level: 1 }, 'Introdução', 10),
        block('heading', { level: 1 }, '   ', 20),
        block('paragraph', { styleId: 'Secao' }, 'Escopo\tdo  trabalho', 30),
      ],
      SHEET,
    )

    expect(entries).toEqual([
      { level: 1, text: 'Introdução', pos: 10 },
      { level: 2, text: 'Escopo do trabalho', pos: 30 },
    ])
  })
})

describe('currentEntryIndex', () => {
  const entries = [
    { level: 1, text: 'A', pos: 10 },
    { level: 2, text: 'B', pos: 30 },
  ]

  it('is the last heading that starts before the position', () => {
    expect(currentEntryIndex(entries, 5)).toBe(-1)
    expect(currentEntryIndex(entries, 10)).toBe(0)
    expect(currentEntryIndex(entries, 29)).toBe(0)
    expect(currentEntryIndex(entries, 400)).toBe(1)
  })
})
