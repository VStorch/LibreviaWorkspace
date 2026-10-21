import { describe, expect, it } from 'vitest'
import {
  BUILTIN_STYLES,
  blockStyleOf,
  listedStyles,
  styleLabelOf,
  type StyleDefinition,
  type StyleSheet,
} from './styles.js'

/** A Portuguese Word document: the id is translated, the internal name is not. */
const portuguese: StyleSheet = {
  defaults: {
    paragraph: {},
    character: { fontFamily: 'Arial, sans-serif', fontSize: '11pt' },
    paragraphStyleId: 'Padro',
    characterStyleId: null,
  },
  styles: {
    Padro: { id: 'Padro', name: 'Normal', type: 'paragraph', qFormat: true, hidden: false, custom: false },
    Ttulo1: {
      id: 'Ttulo1',
      name: 'heading 1',
      type: 'paragraph',
      qFormat: true,
      hidden: false,
      custom: false,
      uiPriority: 9,
    },
    Citao: {
      id: 'Citao',
      name: 'Citação longa',
      type: 'paragraph',
      qFormat: true,
      hidden: false,
      custom: true,
      uiPriority: 1,
    },
    Oculto: {
      id: 'Oculto',
      name: 'Tabela sem nada',
      type: 'paragraph',
      qFormat: false,
      hidden: true,
      custom: false,
    },
  },
}

describe('the style of the block where the cursor is', () => {
  it('is what the block brought from the file', () => {
    const style = blockStyleOf(portuguese, { type: 'paragraph', styleId: 'Citao' })
    expect(style?.name).toBe('Citação longa')
  })

  it('ignores an id the document does not define', () => {
    // In Word a dangling `w:pStyle` is silence: the paragraph comes out as the default. Showing the
    // id anyway would claim the document has a style it does not, which is the case of a `.sdoc`
    // that came from a `.docx` and went back to the minimal package, which does not define
    // `Ttulo1`.
    const style = blockStyleOf(portuguese, { type: 'paragraph', styleId: 'Heading1' })
    expect(style?.id).toBe('Padro')
  })

  it('finds the heading by internal name, not by id', () => {
    // The id is translated and the name is not: the same criterion as the reader and writer
    // (`StyleResolver.HeadingLevelByName`). By id, a heading made on screen would lose its style in
    // every non-English document.
    const style = blockStyleOf(portuguese, { type: 'heading', level: 1 })
    expect(style?.id).toBe('Ttulo1')
  })

  it('falls back to the document default style when the block says nothing', () => {
    const style = blockStyleOf(portuguese, { type: 'paragraph' })
    expect(style?.id).toBe('Padro')
  })

  it('does not invent an answer when the document marks no default style', () => {
    const sheet: StyleSheet = {
      defaults: { paragraph: {}, character: {}, paragraphStyleId: null, characterStyleId: null },
      styles: {},
    }
    expect(blockStyleOf(sheet, { type: 'paragraph' })).toBeNull()
  })

  it("in a new document, the paragraph is Normal and the heading is the level's", () => {
    expect(blockStyleOf(BUILTIN_STYLES, { type: 'paragraph' })?.id).toBe('Normal')
    expect(blockStyleOf(BUILTIN_STYLES, { type: 'heading', level: 3 })?.id).toBe('Heading3')
  })
})

describe('the list the panel shows', () => {
  it('leaves out what the document hides', () => {
    // `w:semiHidden` exists to hide Word's machinery: the default paragraph font and the dozens of
    // table variants every document declares without using. With them, the panel is a list nobody
    // reads.
    expect(listedStyles(portuguese).map((style) => style.id)).toEqual(['Citao', 'Ttulo1', 'Padro'])
  })

  it('sorts by the priority the document declares', () => {
    // Word's order: priority first. `Padro` declares none and goes to the end, not the start, which
    // is where a `?? 0` would put it.
    const priorities = listedStyles(portuguese).map((style) => style.uiPriority ?? 100)
    expect(priorities).toEqual([...priorities].sort((left, right) => left - right))
  })

  it('does not change order between two reads of the same document', () => {
    const twice = listedStyles(BUILTIN_STYLES).map((style) => style.id)
    expect(listedStyles(BUILTIN_STYLES).map((style) => style.id)).toEqual(twice)
  })
})

describe('the style name on screen', () => {
  it('translates the internal name of headings', () => {
    // `heading 1` is the form the file needs for Word to recognize a heading. Showing it would show
    // the file, not the document.
    const heading = BUILTIN_STYLES.styles['Heading1'] as StyleDefinition
    expect(styleLabelOf(heading)).toBe('Título 1')
  })

  it("shows the author's name as written", () => {
    const quotation = portuguese.styles['Citao'] as StyleDefinition
    expect(styleLabelOf(quotation)).toBe('Citação longa')
  })
})
