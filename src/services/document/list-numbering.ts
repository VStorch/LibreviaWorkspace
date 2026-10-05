/**
 * Word counts **per definition** (`w:abstractNum`), not per list: two lists with the same numbering
 * separated by a paragraph continue the count. Levels compose (`%1.%2.`) and each has its own
 * format and start. So the label is computed here and handed over ready, to screen and paper. The
 * definition lives on the list node (`numbering`), as in `ListLevels.cs`.
 */

import { INDENT_STEP_MM } from '@services/units.js'

/** `w:lvl`. */
export interface LevelDef {
  /** `w:numFmt`: `decimal`, `lowerLetter`, `upperRoman`, `bullet`, `none`… */
  readonly fmt: string
  /** `w:lvlText`: `%1.%2.` when numbered; the already translated mark for bullets. */
  readonly text: string
  readonly start: number
  readonly indentMm?: number
  readonly hangingMm?: number
  /** `w:isLgl`: upper levels print in decimal (1.1, not I.a). */
  readonly legal?: boolean
}

export interface NumberingDef {
  /**
   * `a7` is abstract definition 7, which several lists continue; `n12` is `w:num` 12, with its own
   * restart. A list restarted in the editor gets a new key.
   */
  readonly key: string
  readonly abstractId?: number
  readonly levels: readonly LevelDef[]
  /** Level → start value (`w:lvlOverride/w:startOverride`). */
  readonly overrides?: Readonly<Record<string, number>>
}

export const LIST_LEVELS = 9
export const LIST_TYPES: readonly string[] = ['bulletList', 'orderedList']

const HANGING_MM = INDENT_STEP_MM / 2

const round2 = (value: number): number => Math.round(value * 100) / 100

/**
 * Word's levels for a new list: 1. a. i. or • o ▪, cycling every three, half an inch per level. The
 * same as `ListLevels.Defaults`, which writes the list.
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

/** `null` when the attribute has no definition shape. */
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
 * Word letters: after z comes aa, bb, cc, the repeated letter, not the spreadsheet column sequence
 * (aa, ab, ac).
 */
function letter(value: number): string {
  const index = (value - 1) % 26
  return String.fromCharCode(97 + index).repeat(Math.floor((value - 1) / 26) + 1)
}

/**
 * A format the editor does not draw comes out in decimal; the definition goes back to the file
 * intact.
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

/** The ProseMirror tree on screen, the JSON tree in tests. */
export interface ListTreeReader<N> {
  typeOf(node: N): string
  attrsOf(node: N): Readonly<Record<string, unknown>>
  childrenOf(node: N): readonly N[]
}

export interface ListInfo {
  readonly kind: string
  readonly level: number
  readonly key: string
  readonly numId: number | null
  readonly def: NumberingDef
  /** From the margin. */
  readonly indentMm: number
  readonly hangingMm: number
  /** The nested `<ul>` already starts after it. */
  readonly parentIndentMm: number
}

export interface ListNumbering {
  /** One entry per list, in document order (pre-order). */
  readonly lists: ListInfo[]
  /** In document order (pre-order). */
  readonly labels: string[]
  /** Same order as `labels`: the number the HTML and Markdown export put in the list `start`. */
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

/**
 * A single pass in document order: an item depends on everything before it with the same
 * definition.
 */
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
    // Outside a list (a table, a cell), the outer list no longer applies.
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
 * The writer's decision order (`DocxWriter.Flatten`): its own definition; the outer list's when it
 * is the same numbering, or when this one has none and is the same kind; another list's with the
 * same `numId`; and the default, counting alone.
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
    // The `w:num` restart applies the first time the level appears; after that, the item above
    // restarts the level, back to `w:start`.
    const override = started.has(level) ? undefined : list.def.overrides?.[String(level)]
    counts[level] = (override ?? own.start ?? 1) - 1
  }
  started.add(level)
  counts[level] = counts[level]! + 1
  values.push(counts[level])
  // Without `w:lvlRestart`, an item resets the levels below it.
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

/** For each `numId` some list in the document carries. */
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
 * As variables: the node may carry the file's absolute indent, and the rule in `content-styles.ts`
 * uses the relative one, because the nested `<ul>` already starts inside the outer list.
 */
export function listDrawAttrs(info: ListInfo): Record<string, string> {
  const relative = round2(info.indentMm - info.parentIndentMm)
  return {
    'data-list-indent': '',
    style: `--lista-recuo: ${Math.max(0, relative)}mm; --lista-margem: ${Math.min(0, relative)}mm; --lista-pendente: ${Math.max(0, info.hangingMm)}mm`,
  }
}

/**
 * `::before` draws the variable: `attr()` on the inner paragraph would read the wrong attribute.
 */
export function itemDrawAttrs(label: string): Record<string, string> {
  const quoted = label.replace(/["\\]/g, '\\$&').replace(/\n/g, ' ')
  return { 'data-label': label, style: `--lista-marca: "${quoted}"` }
}
