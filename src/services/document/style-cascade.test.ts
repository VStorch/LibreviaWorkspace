import { describe, expect, it } from 'vitest'
import { resolveStyle } from './style-cascade.js'
import { BUILTIN_STYLES, LEGACY_STYLES, StyleType, type StyleSheet } from './styles.js'

describe('resolveStyle', () => {
  it('parte dos padrões do documento e deixa o mais próximo da cadeia com a última palavra', () => {
    const heading = resolveStyle(BUILTIN_STYLES, 'Heading1')
    expect(heading.character).toMatchObject({ fontFamily: 'Calibri', fontSize: '16pt', color: '#2f5496' })
    // O depois do título (0) vence o do padrão (8); a entrelinha, que o título
    // não declara, passa.
    expect(heading.paragraph).toMatchObject({
      spaceBefore: 12,
      spaceAfter: 0,
      lineSpacing: { kind: 'multiple' },
    })
  })

  it('sem id é o estilo padrão de parágrafo', () => {
    expect(resolveStyle(LEGACY_STYLES, null).paragraph).toEqual(LEGACY_STYLES.styles['Normal']?.paragraph)
    expect(resolveStyle(LEGACY_STYLES, null).character).toEqual({
      fontFamily: 'Times New Roman',
      fontSize: '12pt',
    })
  })

  it('id que o documento não define vale só os padrões', () => {
    expect(resolveStyle(BUILTIN_STYLES, 'NaoExiste')).toEqual({
      paragraph: BUILTIN_STYLES.defaults.paragraph,
      character: BUILTIN_STYLES.defaults.character,
    })
  })

  it('não entra em laço num basedOn circular', () => {
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
