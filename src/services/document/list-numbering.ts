/**
 * A numeração das listas, contada como o Word conta.
 *
 * O CSS só sabe contar item a item dentro de um `<ol>`, recomeçando a cada lista.
 * O Word conta de outro jeito, e é esse o que o papel mostra:
 *
 * - **por definição, e não por lista** — duas listas da mesma numeração separadas
 *   por um parágrafo continuam a contagem (1, 2, … 3, 4). No OOXML quem conta é o
 *   `w:abstractNum`; o `w:num` com reinício (`w:startOverride`) conta à parte;
 * - **por nível**, com o texto do nível compondo os números de cima: `%1.%2.`
 *   desenha `2.3.` no terceiro item do segundo;
 * - **com formato por nível** — decimal, letras, romanos, marcador — e o início
 *   que o nível declara (`w:start`).
 *
 * Por isso a marca de cada item é **calculada aqui** e entregue pronta: na tela
 * por decoração, no papel por atributo (ver `extensions/list-numbering.ts`). As
 * duas leem esta mesma conta, que é o que as mantém iguais.
 *
 * A definição mora no nó da lista (`numbering`), como o sidecar a leu — ver
 * `ListLevels.cs`, o espelho desta forma. Nó sem definição (lista criada aqui, ou
 * sublista aberta com Tab) herda a da lista de fora ou recebe a padrão do Word.
 */

/** Um nível da definição: `w:lvl`. */
export interface LevelDef {
  /** `w:numFmt`: `decimal`, `lowerLetter`, `upperRoman`, `bullet`, `none`… */
  readonly fmt: string
  /** `w:lvlText`: `%1.%2.` na numerada; a marca já traduzida na com marcador. */
  readonly text: string
  readonly start: number
  readonly indentMm?: number
  readonly hangingMm?: number
  /** `w:isLgl`: os números de cima saem em decimal (1.1, e não I.a). */
  readonly legal?: boolean
}

/** A definição que o nó da lista leva em `numbering`. */
export interface NumberingDef {
  /**
   * Quem conta: `a7` (a definição abstrata 7, que várias listas continuam) ou
   * `n12` (o `w:num` 12, que tem reinício e conta sozinho). Lista reiniciada no
   * editor ganha uma chave nova até ser gravada.
   */
  readonly key: string
  readonly abstractId?: number
  readonly levels: readonly LevelDef[]
  /** Nível → valor inicial (`w:lvlOverride/w:startOverride`). */
  readonly overrides?: Readonly<Record<string, number>>
}

export const LIST_LEVELS = 9
export const LIST_TYPES: readonly string[] = ['bulletList', 'orderedList']

const INDENT_STEP_MM = 12.7
const HANGING_MM = 6.35

const round2 = (value: number): number => Math.round(value * 100) / 100

/**
 * Os níveis que o Word dá a uma lista nova: 1. a. i. na numerada, • o ▪ na com
 * marcador, repetidos de três em três; recuo de meia polegada por nível e marca
 * pendurada a um quarto.
 *
 * Iguais a `ListLevels.Defaults` no sidecar: é com eles que a lista nova é
 * gravada, e a tela tem de desenhá-la como o arquivo vai sair.
 */
export function defaultLevels(kind: string): LevelDef[] {
  const bullet = kind === 'bulletList'
  const formats = ['decimal', 'lowerLetter', 'lowerRoman']
  const marks = ['•', 'o', '▪']
  return Array.from({ length: LIST_LEVELS }, (_, level) => ({
    fmt: bullet ? 'bullet' : formats[level % 3]!,
    text: bullet ? marks[level % 3]! : `%${level + 1}.`,
    start: 1,
    indentMm: round2(INDENT_STEP_MM * (level + 1)),
    hangingMm: HANGING_MM,
  }))
}

/** A definição do atributo `numbering`, ou `null` quando ele não tem forma de definição. */
export function parseNumbering(value: unknown): NumberingDef | null {
  if (value === null || typeof value !== 'object') return null
  const raw = value as Record<string, unknown>
  if (typeof raw['key'] !== 'string' || !Array.isArray(raw['levels'])) return null
  const levels = (raw['levels'] as unknown[]).filter(
    (level): level is LevelDef =>
      level !== null &&
      typeof level === 'object' &&
      typeof (level as LevelDef).fmt === 'string' &&
      typeof (level as LevelDef).text === 'string',
  )
  if (levels.length === 0) return null
  return value as NumberingDef
}

// --- formatos -----------------------------------------------------------------

function roman(value: number): string {
  const table: Array<[number, string]> = [
    [1000, 'm'],
    [900, 'cm'],
    [500, 'd'],
    [400, 'cd'],
    [100, 'c'],
    [90, 'xc'],
    [50, 'l'],
    [40, 'xl'],
    [10, 'x'],
    [9, 'ix'],
    [5, 'v'],
    [4, 'iv'],
    [1, 'i'],
  ]
  let rest = value
  let text = ''
  for (const [amount, letters] of table) {
    while (rest >= amount) {
      text += letters
      rest -= amount
    }
  }
  return text
}

/**
 * Letra do Word: depois do z vem aa, bb, cc — a letra repetida, e não a
 * sequência de colunas da planilha (aa, ab, ac).
 */
function letter(value: number): string {
  const index = (value - 1) % 26
  return String.fromCharCode(97 + index).repeat(Math.floor((value - 1) / 26) + 1)
}

/**
 * Um número no formato do nível.
 *
 * Formato que o editor não desenha (`ordinal`, `cardinalText`, os de outros
 * alfabetos) sai em decimal: o número certo na forma errada é melhor que a marca
 * sumir. O arquivo não perde nada — a definição volta ao `numbering.xml` intacta.
 */
export function formatNumber(value: number, fmt: string): string {
  if (fmt === 'none' || fmt === 'bullet') return ''
  if (value <= 0 && fmt !== 'decimal' && fmt !== 'decimalZero') return String(value)
  switch (fmt) {
    case 'decimalZero':
      return value >= 0 && value < 10 ? `0${value}` : String(value)
    case 'lowerLetter':
      return letter(value)
    case 'upperLetter':
      return letter(value).toUpperCase()
    case 'lowerRoman':
      return roman(value)
    case 'upperRoman':
      return roman(value).toUpperCase()
    default:
      return String(value)
  }
}

// --- a conta ------------------------------------------------------------------

/** Como ler uma árvore — a do ProseMirror na tela, a do JSON nos testes. */
export interface ListTreeReader<N> {
  typeOf(node: N): string
  attrsOf(node: N): Readonly<Record<string, unknown>>
  childrenOf(node: N): readonly N[]
}

/** O que a conta decide sobre uma lista. */
export interface ListInfo {
  readonly kind: string
  readonly level: number
  readonly key: string
  readonly numId: number | null
  readonly def: NumberingDef
  /** Onde o texto do item começa, a contar da margem. */
  readonly indentMm: number
  readonly hangingMm: number
  /** O recuo da lista de fora — o `<ul>` aninhado já começa depois dele. */
  readonly parentIndentMm: number
}

export interface ListNumbering {
  /** Uma entrada por lista, em ordem de documento (pré-ordem). */
  readonly lists: ListInfo[]
  /** A marca de cada item, em ordem de documento (pré-ordem). */
  readonly labels: string[]
}

function positiveInt(value: unknown): number | null {
  const number = Number(value)
  return value !== null && value !== undefined && Number.isInteger(number) && number > 0 ? number : null
}

function finite(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

interface Counter {
  readonly counts: Array<number | undefined>
  readonly started: Set<number>
}

/**
 * Conta todas as listas de um documento.
 *
 * Um passo só, em ordem de documento, porque é assim que o Word conta: o
 * terceiro item de uma lista depende de tudo o que veio antes dela com a mesma
 * definição, inclusive do que está do outro lado de uma tabela.
 */
export function numberLists<N>(root: N, reader: ListTreeReader<N>): ListNumbering {
  const lists: ListInfo[] = []
  const labels: string[] = []
  const counters = new Map<string, Counter>()
  const byNumId = definitionsByNumId(root, reader)
  let fresh = 0

  const resolve = (node: N, parent: ListInfo | null): ListInfo => {
    const kind = reader.typeOf(node)
    const attrs = reader.attrsOf(node)
    const own = parseNumbering(attrs['numbering'])
    const declared = positiveInt(attrs['numId'])
    const levelAttr = finite(attrs['level'])
    const level = Math.min(
      LIST_LEVELS - 1,
      Math.max(0, levelAttr ?? (parent === null ? 0 : parent.level + 1)),
    )

    // A mesma ordem de decisão do gravador (`DocxWriter.Flatten`): a definição
    // própria; a da lista de fora quando é a mesma numeração, ou quando esta não
    // tem nenhuma e é do mesmo tipo; a de outra lista com o mesmo `numId`; e a
    // padrão, contando sozinha.
    let def: NumberingDef
    let numId = declared
    if (own !== null) {
      def = own
    } else if (
      parent !== null &&
      ((declared !== null && declared === parent.numId) || (declared === null && parent.kind === kind))
    ) {
      def = parent.def
      numId = parent.numId
    } else if (declared !== null && byNumId.has(declared)) {
      def = byNumId.get(declared)!
    } else {
      def = {
        key: declared !== null ? `num${declared}` : `nova${fresh++}`,
        levels: defaultLevels(kind),
      }
    }

    const levelDef = def.levels[level] ?? defaultLevels(kind)[level]!
    const parentIndentMm = parent?.indentMm ?? 0
    return {
      kind,
      level,
      key: def.key,
      numId,
      def,
      indentMm: finite(attrs['indentMm']) ?? levelDef.indentMm ?? round2(INDENT_STEP_MM * (level + 1)),
      hangingMm: finite(attrs['hangingMm']) ?? levelDef.hangingMm ?? 0,
      parentIndentMm,
    }
  }

  const labelOf = (list: ListInfo): string => {
    const counter = counters.get(list.key) ?? { counts: [], started: new Set<number>() }
    counters.set(list.key, counter)
    const { counts, started } = counter
    const level = list.level
    const levels = list.def.levels
    const own = levels[level] ?? defaultLevels(list.kind)[level]!

    if (counts[level] === undefined) {
      // O reinício do `w:num` vale na primeira vez que o nível aparece; depois
      // dela, quem reinicia o nível é o item de cima, e ele volta ao `w:start`.
      const override = started.has(level) ? undefined : list.def.overrides?.[String(level)]
      counts[level] = (override ?? own.start ?? 1) - 1
    }
    started.add(level)
    counts[level] = counts[level]! + 1
    // `w:lvlRestart` ausente: o item de um nível zera os de baixo.
    for (let deeper = level + 1; deeper < LIST_LEVELS; deeper++) counts[deeper] = undefined

    if (own.fmt === 'bullet') return own.text
    if (own.fmt === 'none') return ''

    return own.text.replace(/%([1-9])/g, (_match, digit: string) => {
      const index = Number(digit) - 1
      if (index > level) return ''
      const source = levels[index] ?? own
      const value = counts[index] ?? source.start ?? 1
      const fmt = own.legal === true && index < level ? 'decimal' : source.fmt
      return formatNumber(value, fmt)
    })
  }

  const visit = (node: N, parent: ListInfo | null): void => {
    const type = reader.typeOf(node)
    if (LIST_TYPES.includes(type)) {
      const info = resolve(node, parent)
      lists.push(info)
      for (const child of reader.childrenOf(node)) {
        if (reader.typeOf(child) === 'listItem') {
          labels.push(labelOf(info))
          for (const inner of reader.childrenOf(child)) visit(inner, info)
        } else {
          visit(child, info)
        }
      }
      return
    }
    // Fora de lista (uma tabela, uma célula), a lista de fora não vale mais.
    const context = type === 'listItem' ? parent : null
    for (const child of reader.childrenOf(node)) visit(child, context)
  }

  visit(root, null)
  return { lists, labels }
}

/** A definição de cada `numId` que alguma lista do documento traz. */
function definitionsByNumId<N>(root: N, reader: ListTreeReader<N>): Map<number, NumberingDef> {
  const found = new Map<number, NumberingDef>()
  const walk = (node: N): void => {
    if (LIST_TYPES.includes(reader.typeOf(node))) {
      const attrs = reader.attrsOf(node)
      const numId = positiveInt(attrs['numId'])
      const def = parseNumbering(attrs['numbering'])
      if (numId !== null && def !== null && !found.has(numId)) found.set(numId, def)
    }
    for (const child of reader.childrenOf(node)) walk(child)
  }
  walk(root)
  return found
}

/**
 * Os atributos de desenho de uma lista: o recuo, relativo ao da lista de fora, e
 * a distância da marca.
 *
 * Em variáveis, e não em `padding-left` direto: o nó já pode trazer o recuo do
 * arquivo (`indentMm`), que é absoluto, e a regra em `content-styles.ts` é que
 * decide qual vale — a relativa, porque o `<ul>` aninhado já começa dentro do
 * recuo da lista de fora. Somados, os dois recuos empurravam a sublista para o
 * dobro do que o Word mostra.
 */
export function listDrawAttrs(info: ListInfo): Record<string, string> {
  const relative = round2(info.indentMm - info.parentIndentMm)
  return {
    'data-list-indent': '',
    style: `--lista-recuo: ${Math.max(0, relative)}mm; --lista-margem: ${Math.min(0, relative)}mm; --lista-pendente: ${Math.max(0, info.hangingMm)}mm`,
  }
}

/**
 * O atributo de desenho de um item: a marca pronta.
 *
 * Duas vezes: em `data-label`, para quem lê (testes, acessibilidade), e numa
 * variável, que é o que o `::before` desenha — ele mora no parágrafo de dentro, e
 * `attr()` ali leria o atributo do parágrafo, e não o do item.
 */
export function itemDrawAttrs(label: string): Record<string, string> {
  const quoted = label.replace(/["\\]/g, '\\$&').replace(/\n/g, ' ')
  return { 'data-label': label, style: `--lista-marca: "${quoted}"` }
}
