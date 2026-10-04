/**
 * O MathML mora num atributo do nó, que mora no `.sdoc`, que qualquer um edita
 * à mão: por isso **nunca** passa por `innerHTML`. Um analisador próprio e
 * estrito o lê, e só os elementos do MathML Core e os atributos de apresentação
 * passam; sem `href`, `style` nem evento. Sem DOM, para servir à tela, à
 * exportação e aos testes.
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

/** Comentário, CDATA, DOCTYPE, instrução de processamento e entidade desconhecida recusam o texto inteiro. */
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

/** **Sem** a lista: só para quem a ajeita antes de filtrar (`latex.ts`). */
export function parseMathMl(source: string): MathElement | null {
  return new MathMlParser(source).parse()
}

/** Cada leitura devolve `false` (ou `null`) diante de XML que não é o MathML esperado. */
class MathMlParser {
  private readonly stack: Building[] = []
  private root: Building | null = null
  private count = 0
  private at = 0

  constructor(private readonly source: string) {}

  parse(): MathElement | null {
    while (this.at < this.source.length) {
      if (!this.readNext()) return null
    }
    return this.stack.length === 0 ? this.root : null
  }

  private readNext(): boolean {
    if (this.source[this.at] !== '<') return this.readText()
    return this.source.startsWith('</', this.at) ? this.readClose() : this.readOpen()
  }

  private readText(): boolean {
    const end = this.source.indexOf('<', this.at)
    const raw = this.source.slice(this.at, end === -1 ? this.source.length : end)
    const text = decode(raw)
    if (text === null) return false
    const parent = this.stack.at(-1)
    if (parent === undefined) {
      // Fora do elemento raiz só cabe espaço.
      if (raw.trim() !== '') return false
    } else if (text !== '') {
      parent.children.push(text)
    }
    this.at = end === -1 ? this.source.length : end
    return true
  }

  private readClose(): boolean {
    const match = NAME.exec(this.source.slice(this.at + 2))
    const open = this.stack.pop()
    if (match === null || open === undefined || match[0] !== open.tag) return false
    this.at += 2 + match[0].length
    while (/\s/.test(this.source[this.at] ?? '')) this.at++
    if (this.source[this.at] !== '>') return false
    this.at++
    return true
  }

  private readOpen(): boolean {
    const match = NAME.exec(this.source.slice(this.at + 1))
    // `<!`, `<?` e o que mais não for nome de elemento: recusado.
    if (match === null) return false
    this.at += 1 + match[0].length

    const attrs = this.readAttributes()
    if (attrs === null) return false
    const element: Building = { tag: match[0], attrs, children: [] }
    // Só nomes sem prefixo: o MathML que a conversão escreve usa o namespace padrão.
    if (element.tag.includes(':')) return false
    if (++this.count > MAX_NODES) return false
    if (!this.attach(element)) return false

    if (this.source.startsWith('/>', this.at)) {
      this.at += 2
      return true
    }
    this.at++
    this.stack.push(element)
    return this.stack.length <= MAX_DEPTH
  }

  private attach(element: Building): boolean {
    const parent = this.stack.at(-1)
    if (parent !== undefined) {
      parent.children.push(element)
      return true
    }
    if (this.root !== null) return false
    this.root = element
    return true
  }

  private readAttributes(): Record<string, string> | null {
    const attrs: Record<string, string> = {}
    for (;;) {
      const space = /^\s*/.exec(this.source.slice(this.at))![0].length
      this.at += space
      if (this.source.startsWith('/>', this.at) || this.source[this.at] === '>') return attrs
      if (space === 0) return null
      const attribute = this.readAttribute()
      if (attribute === null) return null
      const [name, value] = attribute
      // O namespace é o do MathML ou nenhum; o resto da declaração cai.
      if (name === 'xmlns' && value !== MATHML_NAMESPACE) return null
      if (!name.startsWith('xmlns')) attrs[name] = value
    }
  }

  private readAttribute(): readonly [string, string] | null {
    const name = NAME.exec(this.source.slice(this.at))
    if (name === null) return null
    this.at += name[0].length
    const equals = /^\s*=\s*/.exec(this.source.slice(this.at))
    if (equals === null) return null
    this.at += equals[0].length
    const quote = this.source[this.at]
    if (quote !== '"' && quote !== "'") return null
    const close = this.source.indexOf(quote, this.at + 1)
    if (close === -1) return null
    const value = decode(this.source.slice(this.at + 1, close))
    if (value === null || value.includes('<')) return null
    this.at = close + 1
    return [name[0], value]
  }
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
