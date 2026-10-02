import { describe, expect, it } from 'vitest'
import { getSchema, getTextBetween, getTextSerializersFromSchema } from '@tiptap/core'
import { DOMParser as ProseMirrorDOMParser, type Node as ProseMirrorNode } from '@tiptap/pm/model'
import { buildEditorExtensions } from '../editor-extensions.js'
import { collectMatches } from './search-replace.js'
import { characterLeaf, textWithoutDeletions } from './track-changes.js'
import { mathAttrsOfMarkup, plainTextOf } from './math.js'

const schema = getSchema(buildEditorExtensions(() => {}))
const FRACTION = '<math><mfrac><mi>a</mi><mi>b</mi></mfrac></math>'

const equation = (attrs: Record<string, unknown>): ProseMirrorNode => schema.nodes['math']!.create(attrs)
const paragraphWith = (...content: ProseMirrorNode[]): ProseMirrorNode =>
  schema.node('doc', null, [schema.node('paragraph', null, content)])

describe('a equação copiada como texto (M11, fase 3)', () => {
  const copied = (node: ProseMirrorNode): string => {
    const doc = paragraphWith(schema.text('a '), node, schema.text(' b'))
    return getTextBetween(
      doc,
      { from: 0, to: doc.content.size },
      {
        textSerializers: getTextSerializersFromSchema(schema),
      },
    )
  }

  it('é o LaTeX guardado', () => {
    expect(copied(equation({ mathml: FRACTION, latex: '\\frac{a}{b}' }))).toBe('a \\frac{a}{b} b')
  })

  it('é o LaTeX tirado do MathML, na que veio de um arquivo', () => {
    expect(plainTextOf(equation({ omml: '<m:oMath/>', mathml: FRACTION, latex: '' }))).toBe('\\frac{a}{b}')
  })

  it('é o marcador quando nem o MathML dá', () => {
    expect(plainTextOf(equation({ mathml: '<script/>', latex: '' }))).toMatch(/^\[equa/)
  })
})

describe('o MathML colado de fora (M11, fase 3)', () => {
  it('vira uma equação nova: sem OMML, filtrada, com o LaTeX tirado dela', () => {
    const attrs = mathAttrsOfMarkup(
      '<math xmlns="http://www.w3.org/1998/Math/MathML" display="block" onclick="x()"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>',
    )
    expect(attrs).toEqual({
      omml: null,
      mathml:
        '<math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>',
      latex: '\\frac{a}{b}',
      display: true,
    })
  })

  it('o que não é MathML não vira equação', () => {
    expect(mathAttrsOfMarkup('<math><mi>x</mi>')).toBe(false)
    expect(mathAttrsOfMarkup('<svg/>')).toBe(false)
  })

  it('as regras de leitura: o embrulho do editor antes do math solto', () => {
    const rules = ProseMirrorDOMParser.fromSchema(schema).rules.map((rule) =>
      'tag' in rule ? rule.tag : null,
    )
    const wrapper = rules.indexOf('span[data-math]')
    expect(wrapper).toBeGreaterThanOrEqual(0)
    expect(rules.indexOf('math')).toBeGreaterThan(wrapper)
  })
})

describe('a equação na busca e na contagem (M11, fase 3)', () => {
  const doc = paragraphWith(
    schema.text('x '),
    equation({ mathml: '<math><mi>x</mi></math>', latex: 'x' }),
    schema.text(' y'),
  )

  it('a busca não acha nada dentro dela, nem atravessa o lugar dela', () => {
    expect(collectMatches(doc, 'x', false)).toEqual([{ from: 1, to: 2 }])
    expect(collectMatches(doc, '  ', false)).toEqual([])
  })

  it('não tem caracteres', () => {
    expect(textWithoutDeletions(doc, undefined, characterLeaf)).toBe('x  y')
  })
})
