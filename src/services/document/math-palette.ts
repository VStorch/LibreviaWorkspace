import type { MessageKey } from '@shared/i18n/index.js'
import { SPECIAL_CHARACTER_GROUPS } from './special-characters.js'
import { latexOfSymbol } from './mathml-latex.js'

/** Greek letters and operators come from `special-characters.ts`, with the same names. */

export interface MathTemplate {
  readonly latex: string
  /** Where the cursor lands after inserting: the first slot to fill. */
  readonly caret: number
  /** The template drawn in characters. */
  readonly label: string
  readonly nameKey: MessageKey
  /** Tests check that Temml reads it. */
  readonly sample: string
}

export interface MathTemplateGroup {
  readonly labelKey: MessageKey
  readonly templates: readonly MathTemplate[]
}

const CARET = '@'

function template(source: string, label: string, nameKey: MessageKey, sample: string): MathTemplate {
  const caret = source.indexOf(CARET)
  return {
    latex: source.replace(CARET, ''),
    caret: caret < 0 ? source.length : caret,
    label,
    nameKey,
    sample,
  }
}

const STRUCTURES: readonly MathTemplate[] = [
  template('\\frac{@}{}', 'a⁄b', 'document.math.template.fraction', '\\frac{a}{b}'),
  template('\\sqrt{@}', '√x', 'document.math.template.sqrt', '\\sqrt{x}'),
  template('\\sqrt[@]{}', 'ⁿ√x', 'document.math.template.root', '\\sqrt[3]{x}'),
  template('{@}^{}', 'xⁿ', 'document.math.template.superscript', 'x^{2}'),
  template('{@}_{}', 'xₙ', 'document.math.template.subscript', 'x_{i}'),
  template('{@}_{}^{}', 'xₙⁿ', 'document.math.template.subSuperscript', 'x_{i}^{2}'),
  template('\\sum_{@}^{}', '∑', 'document.math.template.sum', '\\sum_{i=1}^{n} x_{i}'),
  template('\\int_{@}^{}', '∫', 'document.math.template.integral', '\\int_{0}^{1} f\\,dx'),
  template('\\prod_{@}^{}', '∏', 'document.math.template.product', '\\prod_{k=1}^{n} a_{k}'),
  template(
    '\\begin{pmatrix}@ & \\\\  & \\end{pmatrix}',
    '2×2',
    'document.math.template.matrix2',
    '\\begin{pmatrix}a & b \\\\ c & d\\end{pmatrix}',
  ),
  template(
    '\\begin{pmatrix}@ &  & \\\\  &  & \\\\  &  & \\end{pmatrix}',
    '3×3',
    'document.math.template.matrix3',
    '\\begin{pmatrix}1 & 0 & 0 \\\\ 0 & 1 & 0 \\\\ 0 & 0 & 1\\end{pmatrix}',
  ),
  template('\\left( @ \\right)', '( )', 'document.math.template.parentheses', '\\left( a+b \\right)'),
  template('\\left[ @ \\right]', '[ ]', 'document.math.template.brackets', '\\left[ a+b \\right]'),
  template('\\left\\{ @ \\right\\}', '{ }', 'document.math.template.braces', '\\left\\{ a+b \\right\\}'),
  template('\\left| @ \\right|', '|x|', 'document.math.template.abs', '\\left| x \\right|'),
]

const ACCENTS: readonly MathTemplate[] = [
  template('\\hat{@}', 'x̂', 'document.math.template.hat', '\\hat{x}'),
  template('\\bar{@}', 'x̄', 'document.math.template.bar', '\\bar{x}'),
  template('\\vec{@}', 'v⃗', 'document.math.template.vec', '\\vec{v}'),
  template('\\dot{@}', 'ẋ', 'document.math.template.dot', '\\dot{x}'),
]

const FUNCTIONS: readonly MathTemplate[] = [
  template('\\sin @', 'sin', 'document.math.template.sin', '\\sin x'),
  template('\\cos @', 'cos', 'document.math.template.cos', '\\cos x'),
  template('\\log @', 'log', 'document.math.template.log', '\\log x'),
  template('\\lim_{@} ', 'lim', 'document.math.template.lim', '\\lim_{x \\to 0} f'),
]

/** Operators and relations the character catalog lacks. */
const EXTRA_OPERATORS: readonly { readonly char: string; readonly nameKey: MessageKey }[] = [
  { char: '⋅', nameKey: 'document.math.symbol.cdot' },
  { char: '→', nameKey: 'document.math.symbol.to' },
  { char: '⇒', nameKey: 'document.math.symbol.implies' },
  { char: '≡', nameKey: 'document.math.symbol.equiv' },
  { char: '∝', nameKey: 'document.math.symbol.propto' },
  { char: '∈', nameKey: 'document.math.symbol.in' },
  { char: '⊂', nameKey: 'document.math.symbol.subset' },
  { char: '∪', nameKey: 'document.math.symbol.cup' },
  { char: '∩', nameKey: 'document.math.symbol.cap' },
  { char: '∀', nameKey: 'document.math.symbol.forall' },
  { char: '∃', nameKey: 'document.math.symbol.exists' },
  { char: '∇', nameKey: 'document.math.symbol.nabla' },
]

/** The catalog characters with a LaTeX name, as templates. */
function symbols(
  characters: readonly { readonly char: string; readonly nameKey: MessageKey }[],
): MathTemplate[] {
  return characters.flatMap((character) => {
    const command = latexOfSymbol(character.char)
    if (command === null || !command.startsWith('\\')) return []
    // The space after the command: `\alpha` followed by `x` would be `\alphax`.
    const latex = `${command} `
    return [
      { latex, caret: latex.length, label: character.char, nameKey: character.nameKey, sample: command },
    ]
  })
}

function catalogGroup(
  labelKey: MessageKey,
): readonly { readonly char: string; readonly nameKey: MessageKey }[] {
  return SPECIAL_CHARACTER_GROUPS.find((group) => group.labelKey === labelKey)?.characters ?? []
}

export const MATH_PALETTE: readonly MathTemplateGroup[] = [
  { labelKey: 'document.math.group.structures', templates: STRUCTURES },
  { labelKey: 'document.math.group.accents', templates: ACCENTS },
  { labelKey: 'document.math.group.functions', templates: FUNCTIONS },
  { labelKey: 'chars.group.greek', templates: symbols(catalogGroup('chars.group.greek')) },
  {
    labelKey: 'document.math.group.operators',
    templates: symbols([...catalogGroup('chars.group.math'), ...EXTRA_OPERATORS]),
  },
]
