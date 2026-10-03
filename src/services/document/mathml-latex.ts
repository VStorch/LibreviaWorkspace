import { MATH_BOX_CLASS, sanitizeMathMl, type MathChild, type MathElement } from './mathml.js'

/**
 * O editor de equações precisa de texto para editar a equação vinda do `.docx`.
 * O alvo é o LaTeX que o Temml lê de volta com a mesma estrutura, e não o mais
 * bonito. Entra o MathML já filtrado (`sanitizeMathMl`), do sidecar ou do Temml.
 */

const GREEK: Readonly<Record<string, string>> = {
  α: '\\alpha',
  β: '\\beta',
  γ: '\\gamma',
  δ: '\\delta',
  ϵ: '\\epsilon',
  ε: '\\varepsilon',
  ζ: '\\zeta',
  η: '\\eta',
  θ: '\\theta',
  ϑ: '\\vartheta',
  ι: '\\iota',
  κ: '\\kappa',
  λ: '\\lambda',
  μ: '\\mu',
  ν: '\\nu',
  ξ: '\\xi',
  π: '\\pi',
  ϖ: '\\varpi',
  ρ: '\\rho',
  ϱ: '\\varrho',
  σ: '\\sigma',
  ς: '\\varsigma',
  τ: '\\tau',
  υ: '\\upsilon',
  ϕ: '\\phi',
  φ: '\\varphi',
  χ: '\\chi',
  ψ: '\\psi',
  ω: '\\omega',
  Γ: '\\Gamma',
  Δ: '\\Delta',
  Θ: '\\Theta',
  Λ: '\\Lambda',
  Ξ: '\\Xi',
  Π: '\\Pi',
  Σ: '\\Sigma',
  Υ: '\\Upsilon',
  Φ: '\\Phi',
  Ψ: '\\Psi',
  Ω: '\\Omega',
}

const SYMBOLS: Readonly<Record<string, string>> = {
  '±': '\\pm',
  '∓': '\\mp',
  '×': '\\times',
  '÷': '\\div',
  '⋅': '\\cdot',
  '·': '\\cdot',
  '∗': '\\ast',
  '∘': '\\circ',
  '≤': '\\le',
  '≥': '\\ge',
  '≠': '\\ne',
  '≈': '\\approx',
  '≡': '\\equiv',
  '∼': '\\sim',
  '≃': '\\simeq',
  '≅': '\\cong',
  '∝': '\\propto',
  '≪': '\\ll',
  '≫': '\\gg',
  '→': '\\to',
  '←': '\\leftarrow',
  '↔': '\\leftrightarrow',
  '⇒': '\\Rightarrow',
  '⇐': '\\Leftarrow',
  '⇔': '\\Leftrightarrow',
  '↦': '\\mapsto',
  '∈': '\\in',
  '∉': '\\notin',
  '∋': '\\ni',
  '⊂': '\\subset',
  '⊃': '\\supset',
  '⊆': '\\subseteq',
  '⊇': '\\supseteq',
  '∪': '\\cup',
  '∩': '\\cap',
  '∅': '\\emptyset',
  '∀': '\\forall',
  '∃': '\\exists',
  '¬': '\\neg',
  '∧': '\\wedge',
  '∨': '\\vee',
  '⊕': '\\oplus',
  '⊗': '\\otimes',
  '∥': '\\parallel',
  '⊥': '\\perp',
  '∠': '\\angle',
  '∇': '\\nabla',
  '∂': '\\partial',
  '∞': '\\infty',
  ℏ: '\\hbar',
  ℓ: '\\ell',
  '…': '\\ldots',
  '⋯': '\\cdots',
  '⋮': '\\vdots',
  '⋱': '\\ddots',
  '′': '\\prime',
  '∣': '\\mid',
  '∑': '\\sum',
  '∏': '\\prod',
  '∐': '\\coprod',
  '∫': '\\int',
  '∬': '\\iint',
  '∭': '\\iiint',
  '∮': '\\oint',
  '⋃': '\\bigcup',
  '⋂': '\\bigcap',
  '⋁': '\\bigvee',
  '⋀': '\\bigwedge',
  '−': '-',
  '{': '\\{',
  '}': '\\}',
  '%': '\\%',
  '#': '\\#',
  '&': '\\&',
  $: '\\$',
  _: '\\_',
  '\\': '\\backslash',
  '‖': '\\|',
  '⟨': '\\langle',
  '⟩': '\\rangle',
  '⌊': '\\lfloor',
  '⌋': '\\rfloor',
  '⌈': '\\lceil',
  '⌉': '\\rceil',
}

/** Os nomes de função do LaTeX: `\sin`, `\log`… — o Temml põe o U+2061 depois deles sozinho. */
const FUNCTIONS = new Set([
  'sin',
  'cos',
  'tan',
  'cot',
  'sec',
  'csc',
  'arcsin',
  'arccos',
  'arctan',
  'sinh',
  'cosh',
  'tanh',
  'coth',
  'log',
  'ln',
  'lg',
  'exp',
  'lim',
  'max',
  'min',
  'sup',
  'inf',
  'det',
  'dim',
  'deg',
  'gcd',
  'hom',
  'ker',
  'arg',
  'Pr',
])

/** Os n-ários: o somatório e companhia, que levam os limites por cima e por baixo. */
const NARY = new Set(['∑', '∏', '∐', '∫', '∬', '∭', '∮', '⋃', '⋂', '⋁', '⋀'])
const INTEGRALS = new Set(['∫', '∬', '∭', '∮'])

const ACCENTS: Readonly<Record<string, string>> = {
  '^': '\\hat',
  ˆ: '\\hat',
  '̂': '\\hat',
  '¯': '\\bar',
  '‾': '\\bar',
  '̄': '\\bar',
  '̅': '\\bar',
  '→': '\\vec',
  '⃗': '\\vec',
  '˙': '\\dot',
  '̇': '\\dot',
  '¨': '\\ddot',
  '̈': '\\ddot',
  '~': '\\tilde',
  '˜': '\\tilde',
  '̃': '\\tilde',
  ˇ: '\\check',
  '̌': '\\check',
  '´': '\\acute',
  '́': '\\acute',
  '`': '\\grave',
  '̀': '\\grave',
  '˘': '\\breve',
  '̆': '\\breve',
}

const SPACES: Readonly<Record<string, string>> = {
  '0.1667em': '\\,',
  '0.2222em': '\\:',
  '0.2778em': '\\;',
  '1em': '\\quad',
  '2em': '\\qquad',
  '-0.1667em': '\\!',
}

/** Os ambientes de matriz pelo par de delimitadores em volta. */
const MATRICES: Readonly<Record<string, string>> = {
  '()': 'pmatrix',
  '[]': 'bmatrix',
  '{}': 'Bmatrix',
  '||': 'vmatrix',
  '‖‖': 'Vmatrix',
}

/** Os alfabetos matemáticos do Unicode e o comando de cada um (U+1D400…). */
const ALPHABETS: readonly { readonly start: number; readonly command: string; readonly digits?: boolean }[] =
  [
    { start: 0x1d400, command: '\\mathbf' },
    { start: 0x1d434, command: '\\mathit' },
    { start: 0x1d468, command: '\\boldsymbol' },
    { start: 0x1d49c, command: '\\mathcal' },
    { start: 0x1d504, command: '\\mathfrak' },
    { start: 0x1d538, command: '\\mathbb' },
    { start: 0x1d5a0, command: '\\mathsf' },
    { start: 0x1d670, command: '\\mathtt' },
    { start: 0x1d7ce, command: '\\mathbf', digits: true },
    { start: 0x1d7d8, command: '\\mathbb', digits: true },
    { start: 0x1d7e2, command: '\\mathsf', digits: true },
    { start: 0x1d7f6, command: '\\mathtt', digits: true },
  ]

/** As letras que o Unicode já tinha no bloco Letterlike e deixou fora dos alfabetos. */
const HOLES: Readonly<Record<string, readonly [string, string]>> = {
  ℎ: ['\\mathit', 'h'],
  ℬ: ['\\mathcal', 'B'],
  ℰ: ['\\mathcal', 'E'],
  ℱ: ['\\mathcal', 'F'],
  ℋ: ['\\mathcal', 'H'],
  ℐ: ['\\mathcal', 'I'],
  ℒ: ['\\mathcal', 'L'],
  ℳ: ['\\mathcal', 'M'],
  ℛ: ['\\mathcal', 'R'],
  ℭ: ['\\mathfrak', 'C'],
  ℌ: ['\\mathfrak', 'H'],
  ℑ: ['\\mathfrak', 'I'],
  ℜ: ['\\mathfrak', 'R'],
  ℨ: ['\\mathfrak', 'Z'],
  ℂ: ['\\mathbb', 'C'],
  ℍ: ['\\mathbb', 'H'],
  ℕ: ['\\mathbb', 'N'],
  ℙ: ['\\mathbb', 'P'],
  ℚ: ['\\mathbb', 'Q'],
  ℝ: ['\\mathbb', 'R'],
  ℤ: ['\\mathbb', 'Z'],
}

/** O comando LaTeX de um símbolo (`α` → `\alpha`), ou `null` quando ele não tem nome. */
export function latexOfSymbol(char: string): string | null {
  return GREEK[char] ?? SYMBOLS[char] ?? null
}

export function mathMlToLatex(tree: MathElement): string {
  return row(tree.children)
}

/** O que ela guarda, ou o que sai do MathML na que nunca passou pelo editor. */
export function latexOfEquation(attrs: Readonly<Record<string, unknown>> | undefined): string {
  const latex = typeof attrs?.['latex'] === 'string' ? attrs['latex'].trim() : ''
  if (latex !== '') return latex
  const tree = sanitizeMathMl(typeof attrs?.['mathml'] === 'string' ? attrs['mathml'] : '')
  if (tree === null) return ''
  // O LaTeX que veio junto (o `annotation` do KaTeX, da Wikipédia), quando há.
  const semantics = tree.children.find(
    (child): child is MathElement => typeof child !== 'string' && child.tag === 'semantics',
  )
  const annotation = semantics?.children.find(
    (child): child is MathElement =>
      typeof child !== 'string' &&
      child.tag === 'annotation' &&
      child.attrs['encoding'] === 'application/x-tex',
  )
  const annotated =
    annotation === undefined
      ? ''
      : annotation.children
          .filter((c) => typeof c === 'string')
          .join('')
          .trim()
  return annotated !== '' ? annotated : mathMlToLatex(tree).trim()
}

function elements(children: readonly MathChild[]): MathElement[] {
  return children.filter((child): child is MathElement => typeof child !== 'string')
}

function textOf(node: MathElement): string {
  return node.children.map((child) => (typeof child === 'string' ? child : textOf(child))).join('')
}

/** O `mrow` de um filho só, sem o embrulho. */
function unwrapped(node: MathElement): MathElement {
  let current = node
  while (current.tag === 'mrow' && current.attrs['class'] === undefined) {
    const inner = elements(current.children)
    if (inner.length !== 1) break
    current = inner[0]!
  }
  return current
}

function isApply(node: MathElement | undefined): boolean {
  return node !== undefined && node.tag === 'mo' && textOf(node) === '⁡'
}

function isFence(node: MathElement | undefined): boolean {
  return node !== undefined && node.tag === 'mo' && node.attrs['fence'] === 'true'
}

function isFenced(node: MathElement): boolean {
  if (node.tag !== 'mrow') return false
  const inner = elements(node.children)
  return inner.length >= 2 && (isFence(inner[0]) || isFence(inner.at(-1)))
}

function functionName(node: MathElement): string | null {
  const inner = unwrapped(node)
  if (inner.tag !== 'mi') return null
  const text = textOf(inner)
  return /^[A-Za-z]+$/.test(text) && text.length > 1 ? text : null
}

function naryOperator(node: MathElement): string | null {
  const inner = unwrapped(node)
  if (inner.tag !== 'mo') return null
  const text = textOf(inner).trim()
  return NARY.has(text) ? text : null
}

/**
 * Junta os pedaços: depois de uma palavra de controle (`\alpha`) vem um espaço
 * quando o pedaço seguinte começa com letra, senão `\alphax` seria outro comando.
 */
function join(parts: readonly string[]): string {
  let result = ''
  for (const part of parts) {
    if (part === '') continue
    if (/\\[A-Za-z]+$/.test(result) && /^[A-Za-z]/.test(part)) result += ' '
    result += part
  }
  return result
}

function row(children: readonly MathChild[]): string {
  const items = elements(children)
  const parts: string[] = []
  for (let i = 0; i < items.length; i++) {
    const item = items[i]!
    const previous = items[i - 1]
    const next = items[i + 1]

    if (isApply(item)) continue
    if (item.tag === 'mspace') {
      // O espaço que o Temml põe em volta do nome de função não é da pessoa.
      if (isApply(previous) || (next !== undefined && functionName(next) !== null)) continue
      parts.push(SPACES[item.attrs['width'] ?? ''] ?? '')
      continue
    }

    // O nome seguido do U+2061 é função; o que o LaTeX não conhece por nome
    // (o `sen` do português) vai por `\operatorname`, que traz o U+2061 junto.
    if (isApply(next)) {
      const name = functionName(item)
      if (name !== null) {
        parts.push(FUNCTIONS.has(name) ? `\\${name}` : `\\operatorname{${name}}`)
        continue
      }
    }

    parts.push(node(item))
  }
  return join(parts)
}

function group(node: MathElement | undefined): string {
  if (node === undefined) return '{}'
  const inner =
    node.tag === 'mrow' && !isFenced(node) && node.attrs['class'] === undefined
      ? row(node.children)
      : nodeOf(node)
  return `{${inner}}`
}

/** A base de um índice: sem chaves quando é uma ficha só. */
function base(node: MathElement | undefined): string {
  if (node === undefined) return '{}'
  const written = group(node)
  const inner = written.slice(1, -1)
  return /^(?:[A-Za-z0-9]|\\[A-Za-z]+)$/.test(inner) ? inner : written
}

const nodeOf = (item: MathElement): string => node(item)

function node(item: MathElement): string {
  const args = elements(item.children)
  switch (item.tag) {
    case 'math':
    case 'mstyle':
    case 'mpadded':
    case 'merror':
      return row(item.children)
    case 'semantics':
      return args[0] === undefined ? '' : node(args[0])
    case 'mrow':
      if (item.attrs['class']?.split(/\s+/).includes(MATH_BOX_CLASS) === true)
        return `\\boxed{${row(item.children)}}`
      if (isFenced(item)) return fenced(item)
      return row(item.children)
    case 'mi':
      return identifier(item)
    case 'mn':
      return number(textOf(item))
    case 'mo':
      return operator(textOf(item))
    case 'mtext':
    case 'ms':
      return `\\text{${escapeText(textOf(item))}}`
    case 'mspace':
      return SPACES[item.attrs['width'] ?? ''] ?? ''
    case 'mfrac':
      return /^0(?:\.0*)?(?:px|pt|em)?$/.test(item.attrs['linethickness'] ?? '')
        ? `\\genfrac{}{}{0pt}{}${group(args[0])}${group(args[1])}`
        : `\\frac${group(args[0])}${group(args[1])}`
    case 'msqrt':
      return `\\sqrt{${row(item.children)}}`
    case 'mroot':
      return `\\sqrt[${group(args[1]).slice(1, -1)}]${group(args[0])}`
    case 'msub':
    case 'msup':
    case 'msubsup':
      return scripted(item.tag, args)
    case 'munder':
    case 'mover':
    case 'munderover':
      return limits(item, args)
    case 'mmultiscripts':
      return prescripts(args)
    case 'mtable':
      return `\\begin{matrix}${table(item)}\\end{matrix}`
    case 'mphantom':
      return `\\phantom{${row(item.children)}}`
    case 'none':
    case 'mprescripts':
      return ''
    default:
      return row(item.children)
  }
}

function scripted(tag: 'msub' | 'msup' | 'msubsup', args: readonly MathElement[]): string {
  const [first, second, third] = args
  const nary = first === undefined ? null : naryOperator(first)
  const sub = tag === 'msup' ? null : second
  const sup = tag === 'msub' ? second : tag === 'msup' ? second : third
  const head =
    nary !== null ? `${SYMBOLS[nary] ?? nary}${INTEGRALS.has(nary) ? '' : '\\nolimits'}` : base(first)
  return join([
    head,
    sub === null || sub === undefined ? '' : `_${group(sub)}`,
    tag === 'msub' || sup === undefined ? '' : `^${group(sup)}`,
  ])
}

function limits(item: MathElement, args: readonly MathElement[]): string {
  const [first, second, third] = args
  if (first === undefined) return ''
  const nary = naryOperator(first)
  const name = functionName(first)
  const under = item.tag === 'mover' ? undefined : second
  const over = item.tag === 'munder' ? undefined : item.tag === 'mover' ? second : third

  if (nary !== null || (name !== null && FUNCTIONS.has(name))) {
    const head = nary !== null ? (SYMBOLS[nary] ?? nary) : `\\${name ?? ''}`
    return join([
      head,
      '\\limits',
      under === undefined ? '' : `_${group(under)}`,
      over === undefined ? '' : `^${group(over)}`,
    ])
  }

  if (item.tag === 'munderover') {
    return `\\stackrel${group(over)}{\\underset${group(under)}${group(first)}}`
  }

  const mark = unwrapped(item.tag === 'mover' ? over! : under!)
  const chr = mark.tag === 'mo' ? textOf(mark) : null
  const stretchy = mark.attrs['stretchy'] === 'true'
  const body = group(first)

  if (item.tag === 'munder') {
    if (chr === '⏟' && stretchy) return `\\underbrace${body}`
    if (chr === '‾' && (stretchy || item.attrs['accentunder'] === 'true')) return `\\underline${body}`
    return `\\underset${group(under)}${body}`
  }

  if (chr === '⏞' && stretchy) return `\\overbrace${body}`
  if ((chr === '‾' || chr === '¯') && stretchy) return `\\overline${body}`
  if (chr !== null && !stretchy && ACCENTS[chr] !== undefined) return `${ACCENTS[chr]}${body}`
  // `\stackrel`, e não `\overset`: no Temml este vira `msup`, e o limite por
  // cima (`m:limUpp`) voltaria como expoente.
  return `\\stackrel${group(over)}${body}`
}

function prescripts(args: readonly MathElement[]): string {
  const split = args.findIndex((arg) => arg.tag === 'mprescripts')
  const [first] = args
  if (split < 0) return scripted('msubsup', args)
  const sub = args[split + 1]
  const sup = args[split + 2]
  const subText = sub === undefined || sub.tag === 'none' ? '' : `_${group(sub)}`
  const supText = sup === undefined || sup.tag === 'none' ? '' : `^${group(sup)}`
  return `{}${subText}${supText}${group(first)}`
}

function table(item: MathElement): string {
  return elements(item.children)
    .filter((tr) => tr.tag === 'mtr')
    .map((tr) =>
      elements(tr.children)
        .map((td) => row(td.children))
        .join(' & '),
    )
    .join(' \\\\ ')
}

function delimiter(chr: string, side: 'left' | 'right'): string {
  switch (chr) {
    case '':
      return '.'
    case '(':
    case ')':
    case '[':
    case ']':
    case '|':
    case '/':
      return chr
    case '{':
      return '\\{'
    case '}':
      return '\\}'
    case '‖':
      return '\\|'
    case '⟨':
      return '\\langle'
    case '⟩':
      return '\\rangle'
    case '⌊':
      return '\\lfloor'
    case '⌋':
      return '\\rfloor'
    case '⌈':
      return '\\lceil'
    case '⌉':
      return '\\rceil'
    default:
      return side === 'left' ? '(' : ')'
  }
}

function fenced(item: MathElement): string {
  const children = elements(item.children)
  const opening = isFence(children[0]) ? children[0]! : null
  const closing = children.length > 1 && isFence(children.at(-1)) ? children.at(-1)! : null
  const inner = children.slice(opening === null ? 0 : 1, closing === null ? children.length : -1)
  const open = opening === null ? '' : textOf(opening)
  const close = closing === null ? '' : textOf(closing)
  const stretchy = (opening ?? closing)?.attrs['stretchy'] === 'true'

  const only = inner.length === 1 ? unwrapped(inner[0]!) : null
  const environment = MATRICES[open + close]
  if (only?.tag === 'mtable' && environment !== undefined) {
    return `\\begin{${environment}}${table(only)}\\end{${environment}}`
  }

  if (!stretchy) {
    const plain = (chr: string, side: 'left' | 'right'): string => {
      if (chr === '') return ''
      if (chr === '|') return side === 'left' ? '\\lvert ' : '\\rvert '
      if (chr === '‖') return side === 'left' ? '\\lVert ' : '\\rVert '
      return delimiter(chr, side)
    }
    return join([plain(open, 'left'), row(inner), plain(close, 'right')])
  }

  const parts = inner.map((child) => {
    const text = child.tag === 'mo' ? textOf(child) : ''
    const separator =
      child.tag === 'mo' && (child.attrs['separator'] === 'true' || child.attrs['stretchy'] === 'true')
    if (separator && '|‖/'.includes(text) && text !== '') return `\\middle${delimiter(text, 'left')}`
    return node(child)
  })
  return join([`\\left${delimiter(open, 'left')}`, ...parts, `\\right${delimiter(close, 'right')}`])
}

function styledLetter(code: number): readonly [string, string] | null {
  for (const alphabet of ALPHABETS) {
    const offset = code - alphabet.start
    if (alphabet.digits === true) {
      if (offset >= 0 && offset < 10) return [alphabet.command, String.fromCharCode(48 + offset)]
      continue
    }
    if (offset >= 0 && offset < 52) {
      return [alphabet.command, String.fromCharCode(offset < 26 ? 65 + offset : 97 + offset - 26)]
    }
  }
  return null
}

/** As letras de alfabeto matemático (`𝐯`, `ℝ`) no comando de cada uma: `\mathbf{v}`, `\mathbb{R}`. */
function styled(text: string): string | null {
  const runs: { command: string; letters: string }[] = []
  for (const char of text) {
    const found = HOLES[char] ?? styledLetter(char.codePointAt(0) ?? 0)
    if (found === null) return null
    const [command, letter] = found
    const last = runs.at(-1)
    if (last !== undefined && last.command === command) last.letters += letter
    else runs.push({ command, letters: letter })
  }
  return runs.map((run) => `${run.command}{${run.letters}}`).join('')
}

function identifier(item: MathElement): string {
  const text = textOf(item)
  if (text === '') return ''
  const fromAlphabet = styled(text)
  if (fromAlphabet !== null) return fromAlphabet

  const chars = [...text]
  if (chars.length > 1) {
    if (/^[A-Za-z]+$/.test(text) && FUNCTIONS.has(text)) return `\\${text}`
    return `\\mathrm{${escapeText(text)}}`
  }

  const symbol = GREEK[text] ?? SYMBOLS[text]
  if (symbol !== undefined) return symbol
  if (/^[A-Za-z]$/.test(text) && item.attrs['mathvariant'] === 'normal') return `\\mathrm{${text}}`
  return text
}

function number(text: string): string {
  const fromAlphabet = styled(text.replace(/[.,]/g, ''))
  if (fromAlphabet !== null && /[^\d.,]/.test(text)) return fromAlphabet
  return [...text].map((char) => SYMBOLS[char] ?? char).join('')
}

function operator(text: string): string {
  return join([...text].filter((char) => !'⁡⁢⁣⁤'.includes(char)).map((char) => SYMBOLS[char] ?? char))
}

function escapeText(text: string): string {
  return text.replace(/[\\{}$&#^_%~]/g, (char) => (char === '\\' ? '\\textbackslash ' : `\\${char}`))
}
