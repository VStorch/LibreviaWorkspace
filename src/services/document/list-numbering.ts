/**
 * O Word conta **por definição** (`w:abstractNum`), e não por lista: duas listas
 * da mesma numeração separadas por um parágrafo continuam a contagem. Compõe os
 * níveis (`%1.%2.`) e tem formato e início por nível. Por isso a marca é
 * calculada aqui e entregue pronta, à tela e ao papel. A definição mora no nó da
 * lista (`numbering`), como em `ListLevels.cs`.
 */

import { INDENT_STEP_MM } from '@services/units.js'

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

export interface NumberingDef {
  /**
   * `a7` é a definição abstrata 7, que várias listas continuam; `n12` é o
   * `w:num` 12, com reinício próprio. Lista reiniciada no editor ganha chave nova.
   */
  readonly key: string
  readonly abstractId?: number
  readonly levels: readonly LevelDef[]
  /** Nível → valor inicial (`w:lvlOverride/w:startOverride`). */
  readonly overrides?: Readonly<Record<string, number>>
}

export const LIST_LEVELS = 9
export const LIST_TYPES: readonly string[] = ['bulletList', 'orderedList']

const HANGING_MM = INDENT_STEP_MM / 2

const round2 = (value: number): number => Math.round(value * 100) / 100

/**
 * Os níveis do Word para uma lista nova: 1. a. i. ou • o ▪, de três em três,
 * meia polegada por nível. Iguais a `ListLevels.Defaults`, com que a lista é gravada.
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

/** Formato que o editor não desenha sai em decimal; a definição volta intacta ao arquivo. */
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
  /**
   * O valor do contador no nível de cada item, na mesma ordem de `labels`: o
   * número que a exportação para HTML e Markdown põe no `start` da lista.
   */
  readonly values: number[]
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

/** Um passo só, em ordem de documento: o item depende de tudo o que veio antes com a mesma definição. */
export function numberLists<N>(root: N, reader: ListTreeReader<N>): ListNumbering {
  const lists: ListInfo[] = []
  const labels: string[] = []
  const values: number[] = []
  const counters = new Map<string, Counter>()
  let fresh = 0
  const context: ListContext<N> = {
    reader,
    byNumId: definitionsByNumId(root, reader),
    freshKey: () => `nova${fresh++}`,
  }

  const visit = (node: N, parent: ListInfo | null): void => {
    const type = reader.typeOf(node)
    if (LIST_TYPES.includes(type)) {
      const info = resolveList(node, parent, context)
      lists.push(info)
      for (const child of reader.childrenOf(node)) {
        if (reader.typeOf(child) === 'listItem') {
          labels.push(nextLabel(info, counters, values))
          for (const inner of reader.childrenOf(child)) visit(inner, info)
        } else {
          visit(child, info)
        }
      }
      return
    }
    // Fora de lista (uma tabela, uma célula), a lista de fora não vale mais.
    const outer = type === 'listItem' ? parent : null
    for (const child of reader.childrenOf(node)) visit(child, outer)
  }

  visit(root, null)
  return { lists, labels, values }
}

interface ListContext<N> {
  readonly reader: ListTreeReader<N>
  readonly byNumId: ReadonlyMap<number, NumberingDef>
  readonly freshKey: () => string
}

function resolveList<N>(node: N, parent: ListInfo | null, context: ListContext<N>): ListInfo {
  const kind = context.reader.typeOf(node)
  const attrs = context.reader.attrsOf(node)
  const levelAttr = finite(attrs['level'])
  const level = Math.min(LIST_LEVELS - 1, Math.max(0, levelAttr ?? (parent === null ? 0 : parent.level + 1)))
  const { def, numId } = definitionOf(
    parseNumbering(attrs['numbering']),
    positiveInt(attrs['numId']),
    kind,
    parent,
    context,
  )

  const levelDef = def.levels[level] ?? defaultLevels(kind)[level]!
  return {
    kind,
    level,
    key: def.key,
    numId,
    def,
    indentMm: finite(attrs['indentMm']) ?? levelDef.indentMm ?? round2(INDENT_STEP_MM * (level + 1)),
    hangingMm: finite(attrs['hangingMm']) ?? levelDef.hangingMm ?? 0,
    parentIndentMm: parent?.indentMm ?? 0,
  }
}

/**
 * A mesma ordem de decisão do gravador (`DocxWriter.Flatten`): a definição
 * própria; a da lista de fora quando é a mesma numeração, ou quando esta não tem
 * nenhuma e é do mesmo tipo; a de outra lista com o mesmo `numId`; e a padrão,
 * contando sozinha.
 */
function definitionOf<N>(
  own: NumberingDef | null,
  declared: number | null,
  kind: string,
  parent: ListInfo | null,
  context: ListContext<N>,
): { def: NumberingDef; numId: number | null } {
  if (own !== null) return { def: own, numId: declared }
  if (
    parent !== null &&
    ((declared !== null && declared === parent.numId) || (declared === null && parent.kind === kind))
  ) {
    return { def: parent.def, numId: parent.numId }
  }
  const shared = declared === null ? undefined : context.byNumId.get(declared)
  if (shared !== undefined) return { def: shared, numId: declared }
  const key = declared !== null ? `num${declared}` : context.freshKey()
  return { def: { key, levels: defaultLevels(kind) }, numId: declared }
}

function nextLabel(list: ListInfo, counters: Map<string, Counter>, values: number[]): string {
  const counter = counters.get(list.key) ?? { counts: [], started: new Set<number>() }
  counters.set(list.key, counter)
  const { counts, started } = counter
  const level = list.level
  const own = list.def.levels[level] ?? defaultLevels(list.kind)[level]!

  if (counts[level] === undefined) {
    // O reinício do `w:num` vale na primeira vez que o nível aparece; depois
    // dela, quem reinicia o nível é o item de cima, e ele volta ao `w:start`.
    const override = started.has(level) ? undefined : list.def.overrides?.[String(level)]
    counts[level] = (override ?? own.start ?? 1) - 1
  }
  started.add(level)
  counts[level] = counts[level]! + 1
  values.push(counts[level])
  // `w:lvlRestart` ausente: o item de um nível zera os de baixo.
  for (let deeper = level + 1; deeper < LIST_LEVELS; deeper++) counts[deeper] = undefined

  if (own.fmt === 'bullet') return own.text
  if (own.fmt === 'none') return ''

  return own.text.replace(/%([1-9])/g, (_match, digit: string) => {
    const index = Number(digit) - 1
    if (index > level) return ''
    const source = list.def.levels[index] ?? own
    const value = counts[index] ?? source.start ?? 1
    const fmt = own.legal === true && index < level ? 'decimal' : source.fmt
    return formatNumber(value, fmt)
  })
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
 * Em variáveis: o nó pode trazer o recuo absoluto do arquivo, e a regra em
 * `content-styles.ts` usa o relativo, porque o `<ul>` aninhado já começa dentro
 * da lista de fora.
 */
export function listDrawAttrs(info: ListInfo): Record<string, string> {
  const relative = round2(info.indentMm - info.parentIndentMm)
  return {
    'data-list-indent': '',
    style: `--lista-recuo: ${Math.max(0, relative)}mm; --lista-margem: ${Math.min(0, relative)}mm; --lista-pendente: ${Math.max(0, info.hangingMm)}mm`,
  }
}

/** A variável é o que o `::before` desenha: `attr()` no parágrafo de dentro leria o atributo errado. */
export function itemDrawAttrs(label: string): Record<string, string> {
  const quoted = label.replace(/["\\]/g, '\\$&').replace(/\n/g, ' ')
  return { 'data-label': label, style: `--lista-marca: "${quoted}"` }
}
