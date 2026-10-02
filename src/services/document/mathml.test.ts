import { describe, expect, it } from 'vitest'
import { mathMlToString, mathText, sanitizeMathMl } from './mathml.js'

const NS = 'http://www.w3.org/1998/Math/MathML'

/** O que o sidecar escreve para πr² (ver OmmlMath.cs). */
const AREA = `<math display="inline" xmlns="${NS}"><mrow><mi>π</mi><msup><mrow><mi>r</mi></mrow><mrow><mn>2</mn></mrow></msup></mrow></math>`

describe('o filtro do MathML das equações', () => {
  it('deixa passar o que a conversão do sidecar escreve, sem mudar nada', () => {
    const tree = sanitizeMathMl(AREA)
    expect(tree).not.toBeNull()
    expect(mathMlToString(tree!)).toBe(
      AREA.replace(` display="inline" xmlns="${NS}"`, ` xmlns="${NS}" display="inline"`),
    )
    expect(mathText(tree!)).toBe('πr2')
  })

  it('leva os atributos de apresentação e a classe da caixa', () => {
    const tree = sanitizeMathMl(
      `<math xmlns="${NS}"><mrow class="omml-caixa outra"><mfrac linethickness="0"><mi mathvariant="normal">sin</mi><mo stretchy="true" largeop="true">∑</mo></mfrac></mrow></math>`,
    )
    expect(mathMlToString(tree!)).toBe(
      `<math xmlns="${NS}"><mrow class="omml-caixa"><mfrac linethickness="0"><mi mathvariant="normal">sin</mi><mo stretchy="true" largeop="true">∑</mo></mfrac></mrow></math>`,
    )
  })

  it('derruba elemento de fora da lista com tudo o que tem dentro, e atributo de fora sozinho', () => {
    const tree = sanitizeMathMl(
      `<math xmlns="${NS}"><mi onclick="alert(1)" style="color:red" href="javascript:alert(1)">x</mi>` +
        `<script>alert(1)</script><maction actiontype="toggle"><mi>y</mi></maction>` +
        `<mtext>a</mtext><foreignObject><img src="x" onerror="alert(1)"/></foreignObject></math>`,
    )
    expect(mathMlToString(tree!)).toBe(`<math xmlns="${NS}"><mi>x</mi><mtext>a</mtext></math>`)
  })

  it('escapa o texto e os atributos na volta a texto', () => {
    const tree = sanitizeMathMl(
      `<math xmlns="${NS}"><mo>&lt;</mo><mtext>&amp;&#x3C;img&#62;</mtext><mi dir='"'>a</mi></math>`,
    )
    expect(mathMlToString(tree!)).toBe(
      `<math xmlns="${NS}"><mo>&lt;</mo><mtext>&amp;&lt;img&gt;</mtext><mi dir="&quot;">a</mi></math>`,
    )
  })

  it('recusa o que não é um math bem formado', () => {
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

  it('recusa a árvore funda demais', () => {
    const deep = `<math xmlns="${NS}">${'<mrow>'.repeat(500)}${'</mrow>'.repeat(500)}</math>`
    expect(sanitizeMathMl(deep)).toBeNull()
  })

  it('aceita o elemento vazio fechado em si e o espaço do XML entre filhos', () => {
    const tree = sanitizeMathMl(
      `<math xmlns="${NS}">\n  <mmultiscripts><mi>C</mi><mprescripts /><mn>1</mn><mn>2</mn></mmultiscripts>\n</math>`,
    )
    expect(mathMlToString(tree!)).toBe(
      `<math xmlns="${NS}"><mmultiscripts><mi>C</mi><mprescripts></mprescripts><mn>1</mn><mn>2</mn></mmultiscripts></math>`,
    )
  })
})
