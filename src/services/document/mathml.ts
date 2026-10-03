/**
 * O MathML de uma equação, lido e filtrado antes de chegar à tela.
 *
 * O MathML vem do sidecar (`OmmlMath.cs`), mas mora num atributo do nó — e o nó
 * mora no `.sdoc`, que é um arquivo que qualquer um edita à mão. Por isso ele
 * **nunca** vira HTML por `innerHTML`: é lido aqui por um analisador próprio,
 * estrito, que só reconhece elemento, atributo, texto e as entidades do XML, e
 * passado por uma lista do que pode existir — os elementos do MathML Core que a
 * conversão produz e os atributos de apresentação deles. Elemento fora da lista
 * cai com tudo o que tem dentro; atributo fora da lista cai sozinho. Não há
 * `href`, nem `style`, nem evento: nada que execute ou que busque coisa fora.
 *
 * Sem DOM, de propósito: o mesmo filtro serve à tela (que monta a árvore com
 * `createElementNS`), à exportação em HTML (que a escreve como texto escapado) e
 * aos testes, que não têm DOM neste projeto.
 */

export const MATHML_NAMESPACE = 'http://www.w3.org/1998/Math/MathML'

/** A classe da caixa (`m:borderBox`) — ver `OmmlMath.BoxClass`. */
export const MATH_BOX_CLASS = 'omml-caixa'

export interface MathElement {
  readonly tag: string
  readonly attrs: Readonly<Record<string, string>>
  readonly children: readonly MathChild[]
}

export type MathChild = MathElement | string

const ELEMENTS = new Set([
  'math',
  'mrow',
  'mi',
  'mn',
  'mo',
  'mtext',
  'mspace',
  'ms',
  'mfrac',
  'msqrt',
  'mroot',
  'msub',
  'msup',
  'msubsup',
  'munder',
  'mover',
  'munderover',
  'mmultiscripts',
  'mprescripts',
  'none',
  'mtable',
  'mtr',
  'mtd',
  'mpadded',
  'mphantom',
  'mstyle',
  'merror',
  'semantics',
  'annotation',
])

/** Os elementos que levam texto; nos outros, o texto entre filhos é só espaço do XML. */
const TOKENS = new Set(['mi', 'mn', 'mo', 'mtext', 'ms', 'annotation'])

const ATTRIBUTES = new Set([
  'display',
  'displaystyle',
  'scriptlevel',
  'mathvariant',
  'linethickness',
  'largeop',
  'movablelimits',
  'stretchy',
  'symmetric',
  'fence',
  'separator',
  'accent',
  'accentunder',
  'lspace',
  'rspace',
  'minsize',
  'maxsize',
  'columnalign',
  'rowalign',
  'columnspan',
  'rowspan',
  'width',
  'height',
  'depth',
  'voffset',
  'dir',
  'encoding',
  'class',
])

/** As classes que o CSS do editor conhece; outra qualquer não tem por que chegar. */
const CLASSES = new Set([MATH_BOX_CLASS])

/** Profundidade e tamanho máximos: o que passa disso não é equação, é ataque. */
const MAX_DEPTH = 128
const MAX_NODES = 50_000

/**
 * O MathML filtrado — ou `null` quando o texto não é um `math` bem formado.
 *
 * Bem formado no sentido estrito: comentário, CDATA, DOCTYPE, instrução de
 * processamento e entidade fora das cinco do XML recusam o texto inteiro, em vez
 * de serem adivinhados.
 */
export function sanitizeMathMl(source: string): MathElement | null {
  const parsed = parseMathMl(source)
  if (parsed === null || parsed.tag !== 'math') return null
  return clean(parsed)
}

function clean(element: MathElement): MathElement | null {
  if (!ELEMENTS.has(element.tag)) return null

  const attrs: Record<string, string> = {}
  for (const [name, value] of Object.entries(element.attrs)) {
    if (!ATTRIBUTES.has(name)) continue
    if (name === 'class') {
      const classes = value.split(/\s+/).filter((item) => CLASSES.has(item))
      if (classes.length > 0) attrs[name] = classes.join(' ')
      continue
    }
    attrs[name] = value
  }

  const token = TOKENS.has(element.tag)
  const children: MathChild[] = []
  for (const child of element.children) {
    if (typeof child === 'string') {
      if (token) children.push(child)
      continue
    }
    const kept = clean(child)
    if (kept !== null) children.push(kept)
  }

  return { tag: element.tag, attrs, children }
}

/** O MathML filtrado de volta a texto, escapado — para a exportação em HTML. */
export function mathMlToString(element: MathElement): string {
  const attrs = Object.entries(element.attrs)
    .map(([name, value]) => ` ${name}="${escapeXml(value)}"`)
    .join('')
  const namespace = element.tag === 'math' ? ` xmlns="${MATHML_NAMESPACE}"` : ''
  const inner = element.children
    .map((child) => (typeof child === 'string' ? escapeXml(child) : mathMlToString(child)))
    .join('')
  return `<${element.tag}${namespace}${attrs}>${inner}</${element.tag}>`
}

/** O texto das fichas, na ordem — o que a equação "diz" para quem não a vê. */
export function mathText(element: MathElement): string {
  return element.children
    .map((child) => (typeof child === 'string' ? child : child.tag === 'annotation' ? '' : mathText(child)))
    .join('')
}

function escapeXml(text: string): string {
  return text.replace(/[&<>"]/g, (char) =>
    char === '&' ? '&amp;' : char === '<' ? '&lt;' : char === '>' ? '&gt;' : '&quot;',
  )
}

// --- o analisador ------------------------------------------------------------

const NAME = /^[A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?/
const NAMED_ENTITIES: Readonly<Record<string, string>> = {
  lt: '<',
  gt: '>',
  amp: '&',
  quot: '"',
  apos: "'",
}

interface Building {
  tag: string
  attrs: Record<string, string>
  children: MathChild[]
}

/**
 * A árvore do texto, **sem** a lista: só para quem a ajeita antes de filtrar
 * (`latex.ts`, que traduz o que o Temml escreve fora do MathML Core). Quem
 * desenha usa `sanitizeMathMl`.
 */
export function parseMathMl(source: string): MathElement | null {
  const stack: Building[] = []
  let root: Building | null = null
  let count = 0
  let i = 0

  while (i < source.length) {
    if (source[i] !== '<') {
      const end = source.indexOf('<', i)
      const raw = source.slice(i, end === -1 ? source.length : end)
      const text = decode(raw)
      if (text === null) return null
      const parent = stack.at(-1)
      if (parent === undefined) {
        // Fora do elemento raiz só cabe espaço.
        if (raw.trim() !== '') return null
      } else if (text !== '') {
        parent.children.push(text)
      }
      i = end === -1 ? source.length : end
      continue
    }

    if (source.startsWith('</', i)) {
      const match = NAME.exec(source.slice(i + 2))
      const open = stack.pop()
      if (match === null || open === undefined || match[0] !== open.tag) return null
      i += 2 + match[0].length
      while (/\s/.test(source[i] ?? '')) i++
      if (source[i] !== '>') return null
      i++
      continue
    }

    const match = NAME.exec(source.slice(i + 1))
    // `<!`, `<?` e o que mais não for nome de elemento: recusado.
    if (match === null) return null
    i += 1 + match[0].length

    const attrs: Record<string, string> = {}
    for (;;) {
      const space = /^\s*/.exec(source.slice(i))![0].length
      i += space
      if (source.startsWith('/>', i) || source[i] === '>') break
      if (space === 0) return null
      const name = NAME.exec(source.slice(i))
      if (name === null) return null
      i += name[0].length
      const equals = /^\s*=\s*/.exec(source.slice(i))
      if (equals === null) return null
      i += equals[0].length
      const quote = source[i]
      if (quote !== '"' && quote !== "'") return null
      const close = source.indexOf(quote, i + 1)
      if (close === -1) return null
      const value = decode(source.slice(i + 1, close))
      if (value === null || value.includes('<')) return null
      // O namespace é o do MathML ou nenhum; o resto da declaração cai.
      if (name[0] === 'xmlns' && value !== MATHML_NAMESPACE) return null
      if (!name[0].startsWith('xmlns')) attrs[name[0]] = value
      i = close + 1
    }

    const element: Building = { tag: match[0], attrs, children: [] }
    // Só nomes sem prefixo: o MathML que a conversão escreve usa o namespace padrão.
    if (element.tag.includes(':')) return null
    if (++count > MAX_NODES) return null

    const parent = stack.at(-1)
    if (parent === undefined) {
      if (root !== null) return null
      root = element
    } else {
      parent.children.push(element)
    }

    if (source.startsWith('/>', i)) {
      i += 2
      continue
    }
    i++
    stack.push(element)
    if (stack.length > MAX_DEPTH) return null
  }

  return stack.length === 0 ? root : null
}

/** Texto com as entidades do XML resolvidas — ou `null` diante de uma que não existe. */
function decode(raw: string): string | null {
  let failed = false
  const text = raw.replace(/&(#x[0-9A-Fa-f]+|#[0-9]+|[A-Za-z]+);|&/g, (whole, entity: string | undefined) => {
    if (entity === undefined) {
      failed = true
      return whole
    }
    if (entity.startsWith('#')) {
      const code = entity.startsWith('#x') ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10)
      if (!Number.isFinite(code) || code < 1 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) {
        failed = true
        return whole
      }
      return String.fromCodePoint(code)
    }
    const named = NAMED_ENTITIES[entity]
    if (named === undefined) failed = true
    return named ?? whole
  })
  return failed ? null : text
}
