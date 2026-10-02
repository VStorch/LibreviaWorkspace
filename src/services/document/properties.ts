import type { DocumentProperties } from './model.js'

/**
 * A data no formato do pacote (W3CDTF), sem milissegundos — como o Word grava
 * `dcterms:created` e `dcterms:modified`.
 */
export function w3cdtf(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z')
}

export interface PropertiesStamp {
  /** O nome do autor das preferências — vazio é "não assine". */
  readonly author: string
  readonly now: Date
  /** O documento nunca foi gravado: ganha também o criador e a data de criação. */
  readonly fresh: boolean
  /** O documento mudou desde a última gravação. */
  readonly edited: boolean
}

/**
 * O carimbo de gravação: quem modificou, quando, e a revisão seguinte.
 *
 * Como o Word, mas só quando o documento **mudou**: gravar um arquivo que
 * ninguém editou não muda propriedade nenhuma, e `docProps/` volta ao disco byte
 * a byte. O carimbo vai no modelo, e não no sidecar, porque só quem edita sabe
 * se houve edição.
 *
 * Devolve `current` (a mesma referência) quando não há o que carimbar.
 */
export function stampProperties(
  current: DocumentProperties | undefined,
  stamp: PropertiesStamp,
): DocumentProperties | undefined {
  if (!stamp.edited && !stamp.fresh) return current

  const when = w3cdtf(stamp.now)
  const author = stamp.author.trim()
  const revision = Number.parseInt(current?.revision ?? '', 10)

  return {
    ...current,
    ...(stamp.fresh && current?.created === undefined ? { created: when } : {}),
    ...(stamp.fresh && author !== '' && (current?.creator ?? '') === '' ? { creator: author } : {}),
    modified: when,
    ...(author === '' ? {} : { lastModifiedBy: author }),
    revision: String(Number.isFinite(revision) && revision > 0 ? revision + 1 : 1),
  }
}
