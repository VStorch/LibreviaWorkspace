import type { Band, BandPiece } from './band.js'
import type { SheetPlan } from './paginate.js'
import type { DocumentNode, PageSetup, SectionSetup, SectionStart } from './model.js'

/**
 * As seções do documento (M9), e a herança das faixas entre elas.
 *
 * O modelo guarda a última seção em `page` e as anteriores em `sections`, cada
 * uma só com as faixas que **declara** — é o que o arquivo diz, e é o que deixa
 * editar uma faixa sem ter de achar as cópias dela. Quem desenha precisa da
 * faixa que **vale** em cada seção, e é isto que este módulo resolve: a seção que
 * não declara um tipo usa o da anterior ("Vincular ao anterior" no Word), tipo
 * por tipo — capa, par e padrão herdam cada um por si.
 */

/** As chaves das seis faixas de uma seção. */
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
 * As seções com as faixas herdadas já resolvidas.
 *
 * Da segunda em diante, faixa nula é herança; a primeira não tem de quem herdar,
 * e nula ali é "sem faixa". Uma faixa declarada e vazia **não** herda: é a folha
 * limpa que a pessoa quis, e fica vazia.
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

/** O id de seção que o bloco de primeiro nível encerra, se encerra uma. */
export function sectionBreakOf(
  node: { readonly attrs?: Record<string, unknown> | null } | DocumentNode,
): string | null {
  const value = node.attrs?.['sectionBreak']
  return typeof value === 'string' && value.length > 0 ? value : null
}

/** O pedaço de nó que a procura da marca de seção precisa. */
export interface SectionBlock {
  readonly attrs: Record<string, unknown>
  readonly isTextblock?: boolean
  descendants?: (callback: (node: SectionBlock) => boolean | void) => void
}

/**
 * A marca de seção que o bloco de primeiro nível carrega, se carrega.
 *
 * Quase sempre no próprio parágrafo; num item de lista, no parágrafo de dentro
 * — e aí a seção termina com a lista inteira, que é o bloco que a folha corta.
 */
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
 * A seção de cada bloco de primeiro nível, como índice em `allSections`.
 *
 * O bloco pertence à seção da próxima marca — a dele mesmo ou a de um bloco
 * adiante —, porque a marca **fecha** a seção, como o `w:sectPr` no parágrafo.
 * Depois da última marca vem a seção do corpo. Marca de id desconhecido (um
 * parágrafo colado de outro documento) não é marca: a seção dela não existe.
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

/**
 * A seção começa em folha nova?
 *
 * "Próxima página", par e ímpar sempre; a contínua só quando o papel muda —
 * uma folha não troca de tamanho nem de orientação no meio, e o Word a trata
 * como próxima página.
 */
export function startsNewSheet(section: PageSetup, previous: PageSetup | undefined): boolean {
  if (previous === undefined) return false
  const start = section.start ?? 'nextPage'
  if (start === 'continuous' || start === 'nextColumn') return !samePaper(section, previous)
  return true
}

/** A paridade que a primeira folha da seção exige. */
export function parityOf(section: PageSetup): 'even' | 'odd' | null {
  return section.start === 'evenPage' ? 'even' : section.start === 'oddPage' ? 'odd' : null
}

/** O que uma folha desenhada precisa da seção dela. */
export interface SheetSetup {
  /**
   * A configuração da seção, com `pageNumberStart` apontando o número da
   * primeira folha da seção: é o que deixa `bandForPage` e `pageLabel`, que
   * contam a partir de 1, darem a capa e o número certos em qualquer seção.
   */
  readonly page: PageSetup
  /** A folha, contando da primeira da seção (a partir de 1). */
  readonly inSection: number
}

/**
 * A configuração de cada folha desenhada.
 *
 * A folha em branco que a seção par ou ímpar pede já é da seção nova — o papel
 * e a faixa dela —, mas não é a capa: a primeira folha da seção é a que tem
 * texto.
 */
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
    // A folha em branco não é a capa da seção: leva a faixa comum.
    result.push({
      page: { ...section, pageNumberStart: firstNumber, ...(sheet.blank ? { titlePage: false } : {}) },
      inSection,
    })
  }
  return result
}

/** A última seção e as anteriores, como o modelo as guarda. */
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
 * Parte a seção `index` (em `allSections`) numa quebra nova.
 *
 * Como no Word: a seção de cima é uma cópia da que foi partida — papel,
 * margens, faixas e numeração —, e a de baixo continua sendo ela, agora
 * começando do jeito que a quebra pediu. O `w:type` é da seção que começa.
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
 * Aplica o painel de configuração de página "nesta seção" ou "no documento todo".
 *
 * Na seção editada vale tudo o que o painel diz. No documento todo, as outras
 * recebem o papel, as margens, as distâncias das faixas, o formato do número e a
 * capa distinta — o reinício da numeração fica só onde foi pedido, que é o que
 * faz sentido numa seção e seria estranho em todas. "Pares e ímpares" é do
 * documento inteiro no Word, e vai a todas nos dois casos.
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

/** As chaves das faixas de um lado: cabeçalho ou rodapé. */
const KIND_KEYS = {
  header: ['headerBand', 'firstHeaderBand', 'evenHeaderBand'],
  footer: ['footerBand', 'firstFooterBand', 'evenFooterBand'],
} as const satisfies Record<'header' | 'footer', readonly BandKey[]>

/** O cabeçalho (ou rodapé) da seção vem da anterior: ela não declara nenhum dos três tipos. */
export function isLinkedToPrevious(section: PageSetup, index: number, kind: 'header' | 'footer'): boolean {
  return index > 0 && KIND_KEYS[kind].every((key) => section[key] === null || section[key] === undefined)
}

/**
 * "Vincular ao anterior" ligado ou desligado.
 *
 * Ligado, a seção deixa de declarar as faixas daquele lado, e passa a mostrar as
 * da anterior. Desligado, ela ganha uma cópia das que herdava — com os endereços
 * marcados com a seção dona (`s2~rId5:0:1`, a do corpo é `body`): a gravação cria
 * uma parte nova para a cópia (`SectionWriter.ApplyBands`), e editar uma não muda
 * a outra.
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

/** A cópia da faixa herdada, com os endereços marcados pela seção que passa a ser dona dela. */
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
