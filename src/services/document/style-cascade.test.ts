import { describe, expect, it } from 'vitest'
import { effectiveAttrs, resolveStyle } from './style-cascade.js'
import { BUILTIN_STYLES, LEGACY_STYLES, StyleType, type StyleSheet } from './styles.js'

describe('resolveStyle', () => {
  it('starts from the document defaults and lets the closest in the chain have the last word', () => {
    const heading = resolveStyle(BUILTIN_STYLES, 'Heading1')
    expect(heading.character).toMatchObject({ fontFamily: 'Calibri', fontSize: '16pt', color: '#2f5496' })
    // The heading's space after (0) beats the default's (8); line spacing, which the heading does
    // not declare, passes through.
    expect(heading.paragraph).toMatchObject({
      spaceBefore: 12,
      spaceAfter: 0,
      lineSpacing: { kind: 'multiple' },
    })
  })

  it('without an id it is the default paragraph style', () => {
    expect(resolveStyle(LEGACY_STYLES, null).paragraph).toEqual(LEGACY_STYLES.styles['Normal']?.paragraph)
    expect(resolveStyle(LEGACY_STYLES, null).character).toEqual({
      fontFamily: 'Times New Roman',
      fontSize: '12pt',
    })
  })

  it('an id the document does not define gets only the defaults', () => {
    expect(resolveStyle(BUILTIN_STYLES, 'NaoExiste')).toEqual({
      paragraph: BUILTIN_STYLES.defaults.paragraph,
      character: BUILTIN_STYLES.defaults.character,
    })
  })

  it('does not loop on a circular basedOn', () => {
    const style = { name: 'x', type: StyleType.Paragraph, qFormat: false, hidden: false, custom: true }
    const sheet: StyleSheet = {
      defaults: { paragraph: {}, character: {}, paragraphStyleId: null, characterStyleId: null },
      styles: {
        A: { ...style, id: 'A', basedOn: 'B', paragraph: { spaceAfter: 1 } },
        B: { ...style, id: 'B', basedOn: 'A', paragraph: { spaceAfter: 2, spaceBefore: 3 } },
      },
    }
    expect(resolveStyle(sheet, 'A').paragraph).toEqual({ spaceAfter: 1, spaceBefore: 3 })
  })
})

describe('effectiveAttrs', () => {
  it('a heading without an id takes the `heading N` style, and the direct one wins', () => {
    const attrs = effectiveAttrs(
      { type: 'heading', attrs: { level: 1, spaceAfter: 4, keepNext: null } },
      BUILTIN_STYLES,
    )
    // "Keep with next" is not on the node: it comes from the style, and pagination needs to see it.
    expect(attrs).toMatchObject({ keepNext: true, spaceBefore: 12, spaceAfter: 4, fontSize: '16pt' })
  })

  it("a paragraph takes the declared style, in the node's units", () => {
    const attrs = effectiveAttrs({ type: 'paragraph', attrs: { styleId: 'ListParagraph' } }, BUILTIN_STYLES)
    expect(attrs).toMatchObject({ indentMm: 12.7, spaceAfter: 8, lineHeight: '1.3174' })
  })

  it('list and table, and the editor without styles, come back as they are', () => {
    const list = { type: 'bulletList', attrs: { spaceBefore: 3 } }
    expect(effectiveAttrs(list, BUILTIN_STYLES)).toEqual({ spaceBefore: 3 })
    expect(effectiveAttrs({ type: 'paragraph', attrs: { indent: 0 } }, null)).toEqual({ indent: 0 })
  })
})
