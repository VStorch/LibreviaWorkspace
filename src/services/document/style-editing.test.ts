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

describe('criar e modificar estilos', () => {
  it('o id nasce do nome, sem acento, e não colide', () => {
    expect(newStyleId(BUILTIN_STYLES, 'Citação longa')).toBe('Citacaolonga')
    expect(newStyleId(BUILTIN_STYLES, 'Normal')).toBe('Normal2')
    expect(newStyleId(BUILTIN_STYLES, '***')).toBe('Estilo')
  })

  it('o estilo novo é personalizado e herda tudo do pai', () => {
    const { sheet, id } = createStyle(BUILTIN_STYLES, {
      name: 'Destaque',
      type: StyleType.Paragraph,
      basedOn: 'Normal',
      next: 'Normal',
    })
    expect(sheet.styles[id]).toMatchObject({ custom: true, basedOn: 'Normal', next: 'Normal' })
    expect(sheet.styles[id]?.paragraph).toBeUndefined()
    // Folha nova: a antiga continua como estava.
    expect(BUILTIN_STYLES.styles[id]).toBeUndefined()
  })

  it('só o personalizado muda de nome', () => {
    const renamed = withIdentity(BUILTIN_STYLES, 'Heading1', { name: 'Meu título' })
    expect(renamed.styles['Heading1']?.name).toBe('heading 1')
    expect(headingLevelOfStyle(renamed.styles['Heading1'])).toBe(1)

    const { sheet, id } = createStyle(BUILTIN_STYLES, { name: 'A', type: StyleType.Paragraph })
    expect(withIdentity(sheet, id, { name: 'B' }).styles[id]?.name).toBe('B')
  })

  it('OK sem mudança não mexe no estilo; a mudança grava só o campo mudado', () => {
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
    // A entrelinha não mudou: continua herdada do padrão do documento.
    expect(changed.styles['Heading1']?.paragraph?.lineSpacing).toBeUndefined()
    // E a tela muda junto, pela regra do estilo.
    expect(styleSheetCss(changed)).toContain('font-size: 20pt;')
  })

  it('o estilo seguinte do título é o Normal; sem `next`, o próprio', () => {
    expect(nextStyleIdOf(BUILTIN_STYLES, 'Heading1')).toBe('Normal')
    expect(nextStyleIdOf(BUILTIN_STYLES, null)).toBe('Normal')
    expect(nextStyleIdOf(BUILTIN_STYLES, 'ListParagraph')).toBe('ListParagraph')
  })
})
