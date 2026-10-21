import { describe, expect, it } from 'vitest'
import { styleSheetCss } from './style-css.js'
import {
  createStyle,
  headingLevelOfStyle,
  newStyleId,
  nextStyleIdOf,
  styleDraftOf,
  styleWithDraft,
  withIdentity,
} from './style-editing.js'
import { BUILTIN_STYLES, StyleType } from './styles.js'

describe('creating and modifying styles', () => {
  it('the id comes from the name, without accents, and does not collide', () => {
    expect(newStyleId(BUILTIN_STYLES, 'Citação longa')).toBe('Citacaolonga')
    expect(newStyleId(BUILTIN_STYLES, 'Normal')).toBe('Normal2')
    expect(newStyleId(BUILTIN_STYLES, '***')).toBe('Estilo')
  })

  it('the new style is custom and inherits everything from its parent', () => {
    const { sheet, id } = createStyle(BUILTIN_STYLES, {
      name: 'Destaque',
      type: StyleType.Paragraph,
      basedOn: 'Normal',
      next: 'Normal',
    })
    expect(sheet.styles[id]).toMatchObject({ custom: true, basedOn: 'Normal', next: 'Normal' })
    expect(sheet.styles[id]?.paragraph).toBeUndefined()
    // A new sheet: the old one stays as it was.
    expect(BUILTIN_STYLES.styles[id]).toBeUndefined()
  })

  it('only the custom one changes name', () => {
    const renamed = withIdentity(BUILTIN_STYLES, 'Heading1', { name: 'Meu título' })
    expect(renamed.styles['Heading1']?.name).toBe('heading 1')
    expect(headingLevelOfStyle(renamed.styles['Heading1'])).toBe(1)

    const { sheet, id } = createStyle(BUILTIN_STYLES, { name: 'A', type: StyleType.Paragraph })
    expect(withIdentity(sheet, id, { name: 'B' }).styles[id]?.name).toBe('B')
  })

  it('OK with no change leaves the style alone; a change saves only the changed field', () => {
    const draft = styleDraftOf(BUILTIN_STYLES, 'Heading1')
    expect(styleWithDraft(BUILTIN_STYLES, 'Heading1', draft).styles['Heading1']).toEqual(
      BUILTIN_STYLES.styles['Heading1'],
    )

    const changed = styleWithDraft(BUILTIN_STYLES, 'Heading1', {
      ...draft,
      fontSize: 20,
      paragraph: { ...draft.paragraph, spaceAfter: 6 },
    })
    expect(changed.styles['Heading1']?.character?.fontSize).toBe('20pt')
    expect(changed.styles['Heading1']?.paragraph?.spaceAfter).toBe(6)
    // The line spacing did not change: it is still inherited from the document default.
    expect(changed.styles['Heading1']?.paragraph?.lineSpacing).toBeUndefined()
    // And the screen changes along, through the style rule.
    expect(styleSheetCss(changed)).toContain('font-size: 20pt;')
  })

  it("a heading's next style is Normal; without `next`, itself", () => {
    expect(nextStyleIdOf(BUILTIN_STYLES, 'Heading1')).toBe('Normal')
    expect(nextStyleIdOf(BUILTIN_STYLES, null)).toBe('Normal')
    expect(nextStyleIdOf(BUILTIN_STYLES, 'ListParagraph')).toBe('ListParagraph')
  })
})
