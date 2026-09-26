import { describe, expect, it } from 'vitest'
import { ensureCaptionStyle, ensureTocHeadingStyle, ensureTocStyle } from './reference-styles.js'
import { BUILTIN_STYLES, StyleType, type StyleSheet } from './styles.js'

describe('os estilos das referências', () => {
  it('usa o estilo que o documento já tem pelo nome, ainda que com outro id', () => {
    const sheet: StyleSheet = {
      ...BUILTIN_STYLES,
      styles: {
        ...BUILTIN_STYLES.styles,
        Sumrio1: {
          id: 'Sumrio1',
          name: 'toc 1',
          type: StyleType.Paragraph,
          qFormat: false,
          hidden: false,
          custom: false,
        },
      },
    }
    const ensured = ensureTocStyle(sheet, 1)
    expect(ensured.id).toBe('Sumrio1')
    expect(ensured.sheet).toBe(sheet)
  })

  it('cria o que falta com as medidas do Word', () => {
    const second = ensureTocStyle(BUILTIN_STYLES, 2)
    expect(second.id).toBe('TOC2')
    expect(second.sheet.styles['TOC2']).toMatchObject({
      name: 'toc 2',
      paragraph: { indentMm: 3.88, spaceAfter: 5 },
    })

    const heading = ensureTocHeadingStyle(BUILTIN_STYLES)
    expect(heading.sheet.styles[heading.id]).toMatchObject({
      name: 'TOC Heading',
      basedOn: 'Heading1',
      paragraph: { outlineLevel: 9 },
    })

    const caption = ensureCaptionStyle(BUILTIN_STYLES)
    expect(caption.sheet.styles[caption.id]).toMatchObject({ name: 'caption', character: { italic: true } })
  })

  it('não toma o id de um estilo de outro nome', () => {
    const sheet: StyleSheet = {
      ...BUILTIN_STYLES,
      styles: {
        ...BUILTIN_STYLES.styles,
        Caption: {
          id: 'Caption',
          name: 'Minha legenda',
          type: StyleType.Paragraph,
          qFormat: false,
          hidden: false,
          custom: true,
        },
      },
    }
    expect(ensureCaptionStyle(sheet).id).toBe('Caption_2')
  })
})
