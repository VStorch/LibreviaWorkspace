import { describe, expect, it } from 'vitest'
import { currentEntryIndex, outlineLevelOf, outlineOf, type OutlineBlock } from './outline.js'
import { BUILTIN_STYLES, StyleType, type StyleSheet } from './styles.js'

/** A folha do Word com um estilo do autor que é título, e o `TOC Heading` que desliga o nível. */
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
  it('o título vale pelo nível que traz', () => {
    expect(outlineLevelOf({ type: 'heading', attrs: { level: 2 } }, null)).toBe(2)
  })

  it('o parágrafo vale pelo nível que a cascata de estilos lhe dá', () => {
    expect(outlineLevelOf({ type: 'paragraph', attrs: { styleId: 'Capitulo' } }, SHEET)).toBe(1)
    // O nível do próprio estilo vence o herdado.
    expect(outlineLevelOf({ type: 'paragraph', attrs: { styleId: 'Secao' } }, SHEET)).toBe(2)
  })

  it('nível 9 é corpo de texto, ainda que herde de um título', () => {
    expect(outlineLevelOf({ type: 'paragraph', attrs: { styleId: 'TOCHeading' } }, SHEET)).toBeNull()
  })

  it('parágrafo comum, ou sem folha de estilos, não é título', () => {
    expect(outlineLevelOf({ type: 'paragraph', attrs: {} }, SHEET)).toBeNull()
    expect(outlineLevelOf({ type: 'paragraph', attrs: { styleId: 'Capitulo' } }, null)).toBeNull()
  })
})

describe('outlineOf', () => {
  it('lista os títulos na ordem, sem os vazios e com o espaço normalizado', () => {
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

  it('é o último título que começa antes da posição', () => {
    expect(currentEntryIndex(entries, 5)).toBe(-1)
    expect(currentEntryIndex(entries, 10)).toBe(0)
    expect(currentEntryIndex(entries, 29)).toBe(0)
    expect(currentEntryIndex(entries, 400)).toBe(1)
  })
})
