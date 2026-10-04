import type { Band, BandPiece } from './band.js'
import type { SheetPlan } from './paginate.js'
import {
  contentWidthMm,
  type DocumentNode,
  type PageSetup,
  type SectionSetup,
  type SectionStart,
} from './model.js'
import { INDENT_STEP_MM } from '@services/units.js'

/** Meia polegada, o espaço entre colunas que o Word sugere. */
export const DEFAULT_COLUMN_SPACING_MM = INDENT_STEP_MM

/**
 * Cada seção guarda só as faixas que **declara**, como o arquivo. A que não
 * declara um tipo usa o da anterior ("Vincular ao anterior" no Word), tipo por
 * tipo: capa, par e padrão herdam cada um por si.
 */

export const BAND_KEYS = [
  'headerBand',
  'footerBand',
  'firstHeaderBand',
  'firstFooterBand',
  'evenHeaderBand',
  'evenFooterBand',
] as const
export type BandKey = (typeof BAND_KEYS)[number]

/** Todas as seções, em ordem: as anteriores e, por último, a do corpo. */
export function allSections(page: PageSetup, sections: readonly SectionSetup[] = []): PageSetup[] {
  return [...sections, page]
}

/**
 * Da segunda seção em diante, faixa nula herda; na primeira é "sem faixa". Faixa
 * declarada e vazia não herda: é a folha limpa que a pessoa quis.
 */
export function effectiveSections(page: PageSetup, sections: readonly SectionSetup[] = []): PageSetup[] {
  const resolved: PageSetup[] = []
  for (const section of allSections(page, sections)) {
    const previous = resolved.at(-1)
    if (previous === undefined) {
      resolved.push(section)
      continue
    }
    const inherited: Partial<Record<BandKey, Band | null>> = {}
    for (const key of BAND_KEYS) {
      if (section[key] === null || section[key] === undefined) inherited[key] = previous[key]
    }
    resolved.push({ ...section, ...inherited })
  }
  return resolved
}

export function sectionBreakOf(
  node: { readonly attrs?: Record<string, unknown> | null } | DocumentNode,
): string | null {
  const value = node.attrs?.['sectionBreak']
  return typeof value === 'string' && value.length > 0 ? value : null
}

export interface SectionBlock {
  readonly attrs: Record<string, unknown>
  readonly isTextblock?: boolean
  descendants?: (callback: (node: SectionBlock) => boolean | void) => void
}

/** Num item de lista a marca está no parágrafo de dentro, e a seção termina com a lista inteira. */
export function sectionBreakIn(block: SectionBlock): string | null {
  const own = sectionBreakOf(block)
  if (own !== null || block.isTextblock === true || block.descendants === undefined) return own
  let found: string | null = null
  block.descendants((node) => {
    if (found !== null) return false
    found = sectionBreakOf(node)
    return found === null && node.isTextblock !== true
  })
  return found
}

/**
 * Índices em `allSections`. O bloco pertence à seção da próxima marca, porque
 * a marca **fecha** a seção, como o `w:sectPr`. Marca de id desconhecido (de um
 * parágrafo colado de outro documento) não conta.
 */
export function blockSections(
  breaks: readonly (string | null)[],
  sections: readonly SectionSetup[],
): number[] {
  const indexOf = new Map(sections.map((section, index) => [section.id, index]))
  const result = new Array<number>(breaks.length)
  let current = sections.length
  for (let index = breaks.length - 1; index >= 0; index--) {
    const id = breaks[index]
    const known = id === null || id === undefined ? undefined : indexOf.get(id)
    if (known !== undefined) current = known
    result[index] = current
  }
  return result
}

/** O papel e a orientação são os mesmos — a folha não muda de tamanho entre os dois. */
export function samePaper(left: PageSetup, right: PageSetup): boolean {
  return left.size === right.size && left.orientation === right.orientation
}

/** "Próxima página", par e ímpar sempre; a contínua só quando o papel muda, como no Word. */
export function startsNewSheet(section: PageSetup, previous: PageSetup | undefined): boolean {
  if (previous === undefined) return false
  const start = section.start ?? 'nextPage'
  if (start === 'continuous' || start === 'nextColumn') return !samePaper(section, previous)
  return true
}

export function parityOf(section: PageSetup): 'even' | 'odd' | null {
  return section.start === 'evenPage' ? 'even' : section.start === 'oddPage' ? 'odd' : null
}

export interface SheetSetup {
  /** Com `pageNumberStart` no número da primeira folha da seção, para `bandForPage` e `pageLabel` acertarem. */
  readonly page: PageSetup
  /** Contando da primeira folha da seção, a partir de 1. */
  readonly inSection: number
}

/** A folha em branco da seção par ou ímpar já é da seção nova, mas não é a capa. */
export function sheetSetups(
  effective: readonly PageSetup[],
  sheets: readonly Pick<SheetPlan, 'section' | 'number' | 'first' | 'blank'>[],
): SheetSetup[] {
  const result: SheetSetup[] = []
  let inSection = 0
  let firstNumber = 1
  let previous: number | null = null
  for (const sheet of sheets) {
    if (sheet.first || sheet.section !== previous) {
      inSection = 0
      firstNumber = sheet.number
    }
    previous = sheet.section
    inSection += 1
    const section = effective[sheet.section] ?? effective.at(-1)!
    result.push({
      page: { ...section, pageNumberStart: firstNumber, ...(sheet.blank ? { titlePage: false } : {}) },
      inSection,
    })
  }
  return result
}

export interface SectionList {
  readonly page: PageSetup
  readonly sections: readonly SectionSetup[]
}

/** Um id de seção que ainda não existe: `n1`, `n2`… (os lidos do arquivo são `s1`, `s2`…). */
export function freshSectionId(sections: readonly SectionSetup[]): string {
  const taken = new Set(sections.map((section) => section.id))
  let counter = sections.length + 1
  while (taken.has(`n${counter}`)) counter += 1
  return `n${counter}`
}

/**
 * Como no Word: a seção de cima é cópia da partida, e a de baixo continua sendo
 * ela, começando como a quebra pediu. O `w:type` é da seção que começa.
 */
export function withSectionBreak(
  list: SectionList,
  index: number,
  id: string,
  start: SectionStart,
): SectionList {
  const all = allSections(list.page, list.sections)
  const current = all[index] ?? list.page
  // A cópia leva o id novo por cima do da seção partida (a do corpo não tem).
  const upper: SectionSetup = { ...current, id }
  if (index >= list.sections.length) {
    return { page: { ...list.page, start }, sections: [...list.sections, upper] }
  }
  return {
    page: list.page,
    sections: [
      ...list.sections.slice(0, index),
      upper,
      { ...list.sections[index]!, start },
      ...list.sections.slice(index + 1),
    ],
  }
}

/** A lista sem a seção `id` — a quebra dela foi excluída, e o trecho passa à seção de baixo. */
export function withoutSection(list: SectionList, id: string): SectionList {
  return { page: list.page, sections: list.sections.filter((section) => section.id !== id) }
}

/**
 * No documento todo, as outras seções recebem papel, margens, distâncias das
 * faixas, formato do número e capa distinta; o reinício da numeração fica só
 * onde foi pedido. "Pares e ímpares" é do documento inteiro no Word.
 */
export function withPageSetup(
  list: SectionList,
  index: number,
  draft: PageSetup,
  scope: 'section' | 'document',
): SectionList {
  const shared = (section: PageSetup): Partial<PageSetup> =>
    scope === 'document'
      ? {
          size: draft.size,
          orientation: draft.orientation,
          margins: draft.margins,
          headerDistanceMm: draft.headerDistanceMm,
          footerDistanceMm: draft.footerDistanceMm,
          pageNumberFormat: draft.pageNumberFormat,
          titlePage: draft.titlePage,
          evenAndOddHeaders: draft.evenAndOddHeaders,
        }
      : { evenAndOddHeaders: draft.evenAndOddHeaders ?? section.evenAndOddHeaders }
  const last = list.sections.length
  const page = index >= last ? draft : { ...list.page, ...shared(list.page) }
  const sections = list.sections.map((section, at) =>
    at === index ? { ...draft, id: section.id } : { ...section, ...shared(section) },
  )
  return { page, sections }
}

const KIND_KEYS = {
  header: ['headerBand', 'firstHeaderBand', 'evenHeaderBand'],
  footer: ['footerBand', 'firstFooterBand', 'evenFooterBand'],
} as const satisfies Record<'header' | 'footer', readonly BandKey[]>

/** A seção não declara nenhum dos três tipos daquele lado. */
export function isLinkedToPrevious(section: PageSetup, index: number, kind: 'header' | 'footer'): boolean {
  return index > 0 && KIND_KEYS[kind].every((key) => section[key] === null || section[key] === undefined)
}

/**
 * Desligado, a seção ganha uma cópia das faixas que herdava, com endereços
 * marcados pela seção dona (`s2~rId5:0:1`; a do corpo é `body`): a gravação
 * cria uma parte nova para a cópia (`SectionWriter.ApplyBands`).
 */
export function withBandsLinked<T extends PageSetup>(
  section: T,
  inherited: PageSetup | undefined,
  key: string,
  kind: 'header' | 'footer',
  linked: boolean,
): T {
  const changes: Partial<Record<BandKey, Band | null>> = {}
  for (const band of KIND_KEYS[kind]) {
    changes[band] = linked ? null : unlinkedCopy(inherited?.[band] ?? null, key)
  }
  return { ...section, ...changes }
}

function unlinkedCopy(band: Band | null, key: string): Band | null {
  if (band === null) return null
  const mark = (address: string | undefined): string | undefined =>
    address === undefined ? undefined : `${key}~${address.slice(address.indexOf('~') + 1)}`
  const pieces = (list: readonly BandPiece[]): BandPiece[] =>
    list.map((piece) => (piece.pid === undefined ? piece : { ...piece, pid: mark(piece.pid)! }))
  return {
    ...band,
    left: pieces(band.left),
    center: pieces(band.center),
    right: pieces(band.right),
    rows: band.rows.map((row) => ({
      cells: row.cells.map((cell) => ({ ...cell, pieces: pieces(cell.pieces) })),
    })),
    floats: band.floats.map((object) =>
      object.bid === undefined ? object : { ...object, bid: mark(object.bid)! },
    ),
  }
}

export interface ColumnGeometry {
  readonly count: number
  readonly widthMm: number
  /** Da borda esquerda de uma coluna à da seguinte: largura mais espaço. */
  readonly stepMm: number
  readonly spaceMm: number
  readonly separator: boolean
}

/** Larguras diferentes (`widthsMm`) são desenhadas iguais; o arquivo as mantém e o inventário avisa. */
export function columnGeometry(section: PageSetup): ColumnGeometry {
  const count = Math.max(1, Math.round(section.columns?.count ?? 1))
  const spaceMm = count > 1 ? Math.max(0, section.columns?.spaceMm ?? DEFAULT_COLUMN_SPACING_MM) : 0
  const content = contentWidthMm(section)
  const widthMm = Math.max((content - spaceMm * (count - 1)) / count, 1)
  return {
    count,
    widthMm,
    stepMm: widthMm + spaceMm,
    spaceMm,
    separator: section.columns?.separator === true,
  }
}

/** O painel só conhece colunas iguais, e a gravação as iguala (`SectionWriter.ApplyColumns`). */
export function withColumns(
  list: SectionList,
  index: number,
  columns: { readonly count: number; readonly spaceMm: number; readonly separator: boolean },
  scope: 'section' | 'document',
): SectionList {
  const applies = (at: number): boolean => scope === 'document' || at === index
  const last = list.sections.length
  return {
    page: applies(last) ? { ...list.page, columns: { ...columns } } : list.page,
    sections: list.sections.map((section, at) =>
      applies(at) ? { ...section, columns: { ...columns } } : section,
    ),
  }
}

/**
 * A loja guarda uma **biblioteca** de seções por id, e o texto diz quais valem:
 * as marcas na ordem do corpo e o atributo `bodySection` para a última. Assim o
 * desfazer, que só conhece o texto, desfaz também a seção.
 */
export interface ResolvedSections extends SectionList {
  readonly sections: readonly SectionSetup[]
  /** A entrada da biblioteca que faz as vezes da última seção, se há uma. */
  readonly bodyId: string | null
}

function withoutId(section: PageSetup): PageSetup {
  if (!('id' in section)) return section
  const { id: _id, ...rest } = section as SectionSetup
  void _id
  return rest
}

export function resolveSections(
  marks: readonly (string | null)[],
  bodyId: unknown,
  page: PageSetup,
  library: readonly SectionSetup[],
): ResolvedSections {
  const byId = new Map(library.map((section) => [section.id, section]))
  const seen = new Set<string>()
  const sections: SectionSetup[] = []
  for (const id of marks) {
    const section = id === null ? undefined : byId.get(id)
    if (section === undefined || seen.has(section.id)) continue
    seen.add(section.id)
    sections.push(section)
  }
  const body = typeof bodyId === 'string' ? byId.get(bodyId) : undefined
  return {
    page: body === undefined ? page : withoutId(body),
    sections,
    bodyId: body === undefined ? null : body.id,
  }
}

/** Por id e sem apagar nada: a entrada que o texto não usa mais pode voltar com um desfazer. */
export function storeSections(
  next: SectionList,
  bodyId: string | null,
  page: PageSetup,
  library: readonly SectionSetup[],
): { page: PageSetup; library: SectionSetup[] } {
  const changed = new Map(next.sections.map((section) => [section.id, section]))
  if (bodyId !== null) changed.set(bodyId, { ...next.page, id: bodyId })
  const known = new Set(library.map((section) => section.id))
  return {
    page: bodyId === null ? next.page : page,
    library: [
      ...library.map((section) => changed.get(section.id) ?? section),
      ...[...changed.values()].filter((section) => !known.has(section.id)),
    ],
  }
}

/** As marcas de seção de um documento em JSON, bloco a bloco de primeiro nível. */
export function marksOfJson(doc: DocumentNode): (string | null)[] {
  const find = (node: DocumentNode): string | null => {
    const own = sectionBreakOf(node)
    if (own !== null) return own
    for (const child of node.content ?? []) {
      const found = find(child)
      if (found !== null) return found
    }
    return null
  }
  return (doc.content ?? []).map(find)
}

/**
 * Só **acrescenta** à biblioteca, para o desfazer funcionar. A seção de cima é
 * cópia da partida com id novo; a de baixo ganha outra entrada com o começo
 * novo, apontada pela marca que a fecha (`rename`) ou por `bodyId`.
 */
export function planSectionBreak(
  resolved: ResolvedSections,
  library: readonly SectionSetup[],
  index: number,
  start: SectionStart,
): {
  readonly upperId: string
  readonly additions: SectionSetup[]
  readonly rename: { readonly from: string; readonly to: string } | null
  readonly bodyId: string | null
} {
  const all = allSections(resolved.page, resolved.sections)
  const current = withoutId(all[index] ?? resolved.page)
  const upperId = freshSectionId(library)
  const lowerId = freshSectionId([...library, { ...current, id: upperId }])
  const lower = resolved.sections[index]
  return {
    upperId,
    additions: [
      { ...current, id: upperId },
      { ...current, id: lowerId, start },
    ],
    rename: lower === undefined ? null : { from: lower.id, to: lowerId },
    bodyId: lower === undefined ? lowerId : null,
  }
}

/**
 * Ou, na última seção, a quebra que a abre. O trecho de cima passa à seção de
 * baixo, como no Word, e ela recebe as faixas que herdava da excluída.
 */
export function planSectionDelete(
  resolved: ResolvedSections,
  index: number,
): { readonly removeId: string; readonly next: SectionList } | null {
  if (resolved.sections.length === 0) return null
  const target = Math.min(index, resolved.sections.length - 1)
  const removed = resolved.sections[target]!
  const effective = effectiveSections(resolved.page, resolved.sections)
  const lowerIndex = target + 1
  const all = allSections(resolved.page, resolved.sections)
  const lower = all[lowerIndex]!
  const inherited: Partial<Record<BandKey, Band | null>> = {}
  for (const key of BAND_KEYS) {
    if (lower[key] === null || lower[key] === undefined) inherited[key] = effective[target]![key] ?? null
  }
  const updated = { ...lower, ...inherited }
  const sections = resolved.sections.filter((section) => section.id !== removed.id)
  return {
    removeId: removed.id,
    next:
      lowerIndex >= resolved.sections.length
        ? { page: updated, sections }
        : {
            page: resolved.page,
            sections: sections.map((section) =>
              section.id === (lower as SectionSetup).id ? (updated as SectionSetup) : section,
            ),
          },
  }
}
