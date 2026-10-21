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

describe('the equation copied as text (M11, phase 3)', () => {
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

  it('is the stored LaTeX', () => {
    expect(copied(equation({ mathml: FRACTION, latex: '\\frac{a}{b}' }))).toBe('a \\frac{a}{b} b')
  })

  it('is the LaTeX derived from MathML, for one that came from a file', () => {
    expect(plainTextOf(equation({ omml: '<m:oMath/>', mathml: FRACTION, latex: '' }))).toBe('\\frac{a}{b}')
  })

  it('is the placeholder when even MathML does not help', () => {
    expect(plainTextOf(equation({ mathml: '<script/>', latex: '' }))).toMatch(/^\[equa/)
  })
})

describe('MathML pasted from outside (M11, phase 3)', () => {
  it('becomes a new equation: without OMML, filtered, with the LaTeX derived from it', () => {
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

  it('what is not MathML does not become an equation', () => {
    expect(mathAttrsOfMarkup('<math><mi>x</mi>')).toBe(false)
    expect(mathAttrsOfMarkup('<svg/>')).toBe(false)
  })

  it('the parse rules: the editor wrapper comes before bare math', () => {
    const rules = ProseMirrorDOMParser.fromSchema(schema).rules.map((rule) =>
      'tag' in rule ? rule.tag : null,
    )
    const wrapper = rules.indexOf('span[data-math]')
    expect(wrapper).toBeGreaterThanOrEqual(0)
    expect(rules.indexOf('math')).toBeGreaterThan(wrapper)
  })
})

describe('the equation in search and counting (M11, phase 3)', () => {
  const doc = paragraphWith(
    schema.text('x '),
    equation({ mathml: '<math><mi>x</mi></math>', latex: 'x' }),
    schema.text(' y'),
  )

  it('search finds nothing inside it, nor crosses its place', () => {
    expect(collectMatches(doc, 'x', false)).toEqual([{ from: 1, to: 2 }])
    expect(collectMatches(doc, '  ', false)).toEqual([])
  })

  it('has no characters', () => {
    expect(textWithoutDeletions(doc, undefined, characterLeaf)).toBe('x  y')
  })
})
