import type { Band } from './band.js'
import type { DocumentNode, PageSetup, SectionSetup } from './model.js'

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
