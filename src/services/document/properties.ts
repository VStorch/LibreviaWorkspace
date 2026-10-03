import type { DocumentProperties } from './model.js'

/** Sem milissegundos, como o Word grava `dcterms:created` e `dcterms:modified`. */
export function w3cdtf(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z')
}

export interface PropertiesStamp {
  /** Vazio é "não assine". */
  readonly author: string
  readonly now: Date
  /** Nunca gravado: ganha também o criador e a data de criação. */
  readonly fresh: boolean
  readonly edited: boolean
}

/**
 * Como o Word, mas só quando o documento **mudou**: senão `docProps/` volta
 * byte a byte. No modelo, e não no sidecar, porque só quem edita sabe se houve
 * edição. Sem o que carimbar, devolve `current`.
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
