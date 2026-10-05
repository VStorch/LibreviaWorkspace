import { describe, expect, it } from 'vitest'
import { latexToMathMl } from './latex.js'
import { mathMlToLatex } from './mathml-latex.js'
import { MATH_PALETTE } from './math-palette.js'
import { mathMlToString, sanitizeMathMl, type MathChild, type MathElement } from './mathml.js'

/** Or fails the test with Temml's error. */
function render(latex: string, display = false): MathElement {
  const result = latexToMathMl(latex, display)
  if (!result.ok) throw new Error(`${latex}: ${result.error}`)
  return result.tree
}

/**
 * The structure of a MathML tree, to compare two spellings of the same equation: without attributes
 * (except fraction thickness and the box class), without `mspace`, without U+2061 and with simple
 * `mrow`s unwrapped at the top level, since Temml and the sidecar group differently what draws the
 * same.
 */
function shape(node: MathElement): string {
  return items(node.children)
    .map((child) => (typeof child === 'string' ? JSON.stringify(child) : one(child)))
    .join(' ')
}

function one(node: MathElement): string {
  const extra =
    node.tag === 'mfrac' && /^0/.test(node.attrs['linethickness'] ?? '')
      ? '[0]'
      : node.attrs['class'] !== undefined
        ? `[${node.attrs['class']}]`
        : ''
  return `${node.tag}${extra}(${shape(node)})`
}

function items(children: readonly MathChild[]): MathChild[] {
  const result: MathChild[] = []
  for (const child of children) {
    if (typeof child === 'string') {
      // The sidecar circumflex (`^`) and Temml's (`ˆ`) are the same `m:acc`.
      if (child.trim() !== '' && child !== '⁡') result.push(child.trim().replace('^', 'ˆ'))
      continue
    }
    if (child.tag === 'mspace' || (child.tag === 'mo' && child.children.join('') === '⁡')) continue
    if (
      (child.tag === 'mrow' || child.tag === 'mpadded') &&
      child.attrs['class'] === undefined &&
      !fenced(child)
    ) {
      result.push(...items(child.children))
      continue
    }
    result.push(child)
  }
  return result
}

function fenced(node: MathElement): boolean {
  const first = node.children.find((child): child is MathElement => typeof child !== 'string')
  return first?.tag === 'mo' && first.attrs['fence'] === 'true'
}

/** LaTeX → MathML → LaTeX → MathML: the second read has the same structure as the first. */
function expectRoundTrip(tree: MathElement, display = false): string {
  const latex = mathMlToLatex(tree)
  const again = render(latex, display)
  expect(shape(again), latex).toBe(shape(tree))
  return latex
}

const SAMPLES = MATH_PALETTE.flatMap((group) => group.templates.map((template) => template.sample))

describe('LaTeX → MathML (Temml)', () => {
  it.each(SAMPLES)('desenha o modelo da paleta %s', (latex) => {
    const result = latexToMathMl(latex, false)
    expect(result.ok).toBe(true)
  })

  it('o resultado passa pelo filtro sem perder nada', () => {
    for (const latex of SAMPLES) {
      const result = latexToMathMl(latex, true)
      if (!result.ok) throw new Error(latex)
      const again = sanitizeMathMl(result.mathml)
      expect(again).not.toBeNull()
      expect(mathMlToString(again!)).toBe(result.mathml)
      expect(result.mathml).not.toContain('semantics')
      expect(result.mathml).not.toContain('style=')
    }
  })

  it('a fração e a raiz da especificação', () => {
    const tree = render('\\frac{a}{b}+\\sqrt{x}')
    expect(shape(tree)).toBe('mfrac(mi("a") mi("b")) mo("+") msqrt(mi("x"))')
  })

  it('a de exibição leva display="block"', () => {
    const result = latexToMathMl('x', true)
    expect(result.ok && result.tree.attrs['display']).toBe('block')
  })

  it('LaTeX que não fecha devolve o erro do Temml, e não lança', () => {
    for (const latex of ['\\frac{a}', '\\sqrt{', 'a^', '\\naoexiste', '\\left(']) {
      const result = latexToMathMl(latex, false)
      expect(result.ok, latex).toBe(false)
      if (!result.ok) expect(result.error.length).toBeGreaterThan(0)
    }
  })

  it('traduz o que o Temml escreve fora do MathML Core', () => {
    expect(shape(render('\\overline{AB}'))).toBe('mover(mi("A") mi("B") mo("‾"))')
    expect(render('\\overline{AB}').children[0]).toMatchObject({ tag: 'mover', attrs: { accent: 'true' } })
    expect(shape(render('\\boxed{x}'))).toBe('mrow[omml-caixa](mi("x"))')
    expect(shape(render('{}_{a}^{b}X'))).toBe('mmultiscripts(mi("X") mprescripts() mi("a") mi("b"))')
  })
})

describe('MathML → LaTeX', () => {
  it.each(SAMPLES)('ida e volta pela paleta: %s', (latex) => {
    expectRoundTrip(render(latex))
  })

  it.each([
    '\\sum\\limits_{i=1}^{n} x_{i} = \\int_{0}^{\\infty} e^{-t}\\,dt',
    '\\left( x \\middle| y \\right)',
    '\\left\\{ x \\right.',
    '\\binom{n}{k}',
    '\\overbrace{a+b}^{n} \\underbrace{c}_{m}',
    '\\underline{x} \\tilde{y} \\ddot{z} \\check{w}',
    '\\mathbf{v} + \\mathbb{R} + \\mathcal{L} + \\mathrm{d}x',
    '\\alpha \\le \\beta \\ne \\Gamma \\cdot \\Delta',
    '\\operatorname{sen} x + \\lim\\limits_{n \\to \\infty} a_{n}',
    'a \\text{ se } b',
    '{}^{14}_{6}C',
    "f'(x) = \\lvert x \\rvert",
    '\\begin{bmatrix}a \\\\ b\\end{bmatrix} \\begin{vmatrix}1 & 2 \\\\ 3 & 4\\end{vmatrix}',
  ])('ida e volta: %s', (latex) => {
    expectRoundTrip(render(latex))
    expectRoundTrip(render(latex, true), true)
  })

  it('a especificação: \\frac{a}{b}+\\sqrt{x}', () => {
    expect(mathMlToLatex(render('\\frac{a}{b}+\\sqrt{x}'))).toBe('\\frac{a}{b}+\\sqrt{x}')
  })

  /**
   * The MathML the sidecar writes (`OmmlMath.Convert`) for each OMML construct: the LaTeX derived
   * from it, read by Temml, gives the same structure.
   */
  it.each([
    ['m:f', '<mfrac><mrow><mi>a</mi></mrow><mrow><mi>b</mi></mrow></mfrac>'],
    ['m:f noBar', '<mfrac linethickness="0"><mrow><mi>n</mi></mrow><mrow><mi>k</mi></mrow></mfrac>'],
    [
      'm:rad',
      '<msqrt><mrow><mi>x</mi></mrow></msqrt><mroot><mrow><mi>x</mi></mrow><mrow><mn>3</mn></mrow></mroot>',
    ],
    ['m:sSub', '<msub><mrow><mi>x</mi></mrow><mrow><mi>i</mi></mrow></msub>'],
    ['m:sSup', '<msup><mrow><mi>π</mi></mrow><mrow><mn>2</mn></mrow></msup>'],
    ['m:sSubSup', '<msubsup><mrow><mi>x</mi></mrow><mrow><mi>i</mi></mrow><mrow><mn>2</mn></mrow></msubsup>'],
    [
      'm:sPre',
      '<mmultiscripts><mrow><mi>X</mi></mrow><mprescripts/><mrow><mi>a</mi></mrow><mrow><mi>b</mi></mrow></mmultiscripts>',
    ],
    [
      'm:nary',
      '<mrow><munderover><mo largeop="true">∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mrow><mi>n</mi></mrow></munderover><mrow><msub><mrow><mi>x</mi></mrow><mrow><mi>i</mi></mrow></msub></mrow></mrow>',
    ],
    [
      'm:nary subSup',
      '<mrow><msubsup><mo largeop="true" movablelimits="false">∫</mo><mrow><mn>0</mn></mrow><mrow><mn>1</mn></mrow></msubsup><mrow><mi>f</mi></mrow></mrow>',
    ],
    [
      'm:d',
      '<mrow><mo fence="true" stretchy="true">[</mo><mrow><mi>y</mi></mrow><mo separator="true">|</mo><mrow><mi>z</mi></mrow><mo fence="true" stretchy="true">]</mo></mrow>',
    ],
    [
      'm:m',
      '<mrow><mo fence="true" stretchy="true">(</mo><mrow><mtable><mtr><mtd><mrow><mn>1</mn></mrow></mtd><mtd><mrow><mn>0</mn></mrow></mtd></mtr><mtr><mtd><mrow><mn>0</mn></mrow></mtd><mtd><mrow><mn>1</mn></mrow></mtd></mtr></mtable></mrow><mo fence="true" stretchy="true">)</mo></mrow>',
    ],
    [
      'm:eqArr',
      '<mtable><mtr><mtd><mrow><mi>a</mi><mo>=</mo><mn>1</mn></mrow></mtd></mtr><mtr><mtd><mrow><mi>b</mi></mrow></mtd></mtr></mtable>',
    ],
    [
      'm:acc',
      '<mover accent="true"><mrow><mi>v</mi></mrow><mo stretchy="false">→</mo></mover><mover accent="true"><mrow><mi>x</mi></mrow><mo stretchy="false">^</mo></mover>',
    ],
    ['m:bar', '<mover accent="true"><mrow><mi>A</mi><mi>B</mi></mrow><mo stretchy="true">‾</mo></mover>'],
    ['m:groupChr', '<munder><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo stretchy="true">⏟</mo></munder>'],
    [
      'm:limLow',
      '<munder><mrow><mi mathvariant="normal">lim</mi></mrow><mrow><mi>n</mi><mo>→</mo><mi>∞</mi></mrow></munder>',
    ],
    ['m:limUpp', '<mover><mrow><mi>x</mi></mrow><mrow><mn>2</mn></mrow></mover>'],
    [
      'm:func',
      '<mrow><mrow><mi mathvariant="normal">sin</mi></mrow><mo>⁡</mo><mrow><mi>x</mi></mrow></mrow>',
    ],
    [
      'm:func sen',
      '<mrow><mrow><mi mathvariant="normal">sen</mi></mrow><mo>⁡</mo><mrow><mi>x</mi></mrow></mrow>',
    ],
    ['m:borderBox', '<mrow class="omml-caixa"><mrow><mi>z</mi></mrow></mrow>'],
    ['m:r', '<mi>𝐯</mi><mi mathvariant="normal">d</mi><mtext>texto</mtext><mi>Δ</mi><mn>2</mn><mi>a</mi>'],
  ])('ida e volta pelo MathML do sidecar: %s', (_label, inner) => {
    const tree = sanitizeMathMl(`<math display="inline">${inner}</math>`)
    expect(tree).not.toBeNull()
    expectRoundTrip(tree!)
  })
})
