import { describe, expect, it } from 'vitest'
import { mathMlToString, mathText, sanitizeMathMl } from './mathml.js'
import { DOCUMENT_CONTENT_CSS, PRINT_ONLY_CSS } from './content-styles.js'
import { latexOfEquation } from './mathml-latex.js'

const NS = 'http://www.w3.org/1998/Math/MathML'

/** What the sidecar writes for πr² (see OmmlMath.cs). */
const AREA = `<math display="inline" xmlns="${NS}"><mrow><mi>π</mi><msup><mrow><mi>r</mi></mrow><mrow><mn>2</mn></mrow></msup></mrow></math>`

describe("the equations' MathML filter", () => {
  it('lets through what the sidecar conversion writes, unchanged', () => {
    const tree = sanitizeMathMl(AREA)
    expect(tree).not.toBeNull()
    expect(mathMlToString(tree!)).toBe(
      AREA.replace(` display="inline" xmlns="${NS}"`, ` xmlns="${NS}" display="inline"`),
    )
    expect(mathText(tree!)).toBe('πr2')
  })

  it('carries the presentation attributes and the box class', () => {
    const tree = sanitizeMathMl(
      `<math xmlns="${NS}"><mrow class="omml-box other"><mfrac linethickness="0"><mi mathvariant="normal">sin</mi><mo stretchy="true" largeop="true">∑</mo></mfrac></mrow></math>`,
    )
    expect(mathMlToString(tree!)).toBe(
      `<math xmlns="${NS}"><mrow class="omml-box"><mfrac linethickness="0"><mi mathvariant="normal">sin</mi><mo stretchy="true" largeop="true">∑</mo></mfrac></mrow></math>`,
    )
  })

  it('drops an element outside the list with everything inside it, and an outside attribute alone', () => {
    const tree = sanitizeMathMl(
      `<math xmlns="${NS}"><mi onclick="alert(1)" style="color:red" href="javascript:alert(1)">x</mi>` +
        `<script>alert(1)</script><maction actiontype="toggle"><mi>y</mi></maction>` +
        `<mtext>a</mtext><foreignObject><img src="x" onerror="alert(1)"/></foreignObject></math>`,
    )
    expect(mathMlToString(tree!)).toBe(`<math xmlns="${NS}"><mi>x</mi><mtext>a</mtext></math>`)
  })

  it('escapes the text and attributes on the way back to text', () => {
    const tree = sanitizeMathMl(
      `<math xmlns="${NS}"><mo>&lt;</mo><mtext>&amp;&#x3C;img&#62;</mtext><mi dir='"'>a</mi></math>`,
    )
    expect(mathMlToString(tree!)).toBe(
      `<math xmlns="${NS}"><mo>&lt;</mo><mtext>&amp;&lt;img&gt;</mtext><mi dir="&quot;">a</mi></math>`,
    )
  })

  it('rejects what is not a well-formed math', () => {
    const refused = [
      '',
      'texto solto',
      `<mrow xmlns="${NS}"><mi>x</mi></mrow>`,
      `<math xmlns="${NS}"><mi>x</mi>`,
      `<math xmlns="${NS}"><mi>x</mo></math>`,
      `<math xmlns="${NS}"><!-- c --><mi>x</mi></math>`,
      `<math xmlns="${NS}"><![CDATA[<script>]]></math>`,
      `<!DOCTYPE math><math xmlns="${NS}"/>`,
      `<math xmlns="${NS}"><mi>&nbsp;</mi></math>`,
      `<math xmlns="${NS}"><mi>&</mi></math>`,
      `<math xmlns="http://www.w3.org/1999/xhtml"><mi>x</mi></math>`,
      `<math xmlns="${NS}"><html:script xmlns:html="http://www.w3.org/1999/xhtml"/></math>`,
      `<math xmlns="${NS}"/><math xmlns="${NS}"/>`,
      `<math xmlns="${NS}"><mi a=b>x</mi></math>`,
    ]
    for (const source of refused) expect(sanitizeMathMl(source), source).toBeNull()
  })

  it('rejects a tree that is too deep', () => {
    const deep = `<math xmlns="${NS}">${'<mrow>'.repeat(500)}${'</mrow>'.repeat(500)}</math>`
    expect(sanitizeMathMl(deep)).toBeNull()
  })

  it('accepts the self-closed empty element and XML whitespace between children', () => {
    const tree = sanitizeMathMl(
      `<math xmlns="${NS}">\n  <mmultiscripts><mi>C</mi><mprescripts /><mn>1</mn><mn>2</mn></mmultiscripts>\n</math>`,
    )
    expect(mathMlToString(tree!)).toBe(
      `<math xmlns="${NS}"><mmultiscripts><mi>C</mi><mprescripts></mprescripts><mn>1</mn><mn>2</mn></mmultiscripts></math>`,
    )
  })
})

describe('the equation on paper (M11, phase 3)', () => {
  it('the display one does not split across pages, and printing carries the screen CSS', () => {
    expect(DOCUMENT_CONTENT_CSS).toMatch(/\.equation--display \{[^}]*break-inside: avoid/)
    expect(DOCUMENT_CONTENT_CSS).toContain('.page__content .equation math {')
    expect(PRINT_ONLY_CSS).not.toContain('equation')
  })
})

describe('latexOfEquation', () => {
  it('the stored one; otherwise, the TeX annotation; otherwise, the one derived from MathML', () => {
    expect(latexOfEquation({ latex: ' x^2 ', mathml: '<math><mi>y</mi></math>' })).toBe('x^2')
    expect(
      latexOfEquation({
        latex: '',
        mathml:
          '<math><semantics><mrow><mi>y</mi></mrow><annotation encoding="application/x-tex">y_{0}</annotation></semantics></math>',
      }),
    ).toBe('y_{0}')
    expect(latexOfEquation({ mathml: '<math><msup><mi>r</mi><mn>2</mn></msup></math>' })).toBe('r^{2}')
    expect(latexOfEquation({ mathml: '<script/>' })).toBe('')
    expect(latexOfEquation(undefined)).toBe('')
  })
})
