import type { DocumentProperties } from './model.js'

/** Without milliseconds, as Word writes `dcterms:created` and `dcterms:modified`. */
export function w3cdtf(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z')
}

export interface PropertiesStamp {
  /** Empty means "do not sign". */
  readonly author: string
  readonly now: Date
  /** Never saved: also gets the creator and the creation date. */
  readonly fresh: boolean
  readonly edited: boolean
}

/**
 * Like Word, but only when the document **changed**: otherwise `docProps/` goes back byte for byte.
 * In the model, not the sidecar, because only the editor knows whether there was an edit. With
 * nothing to stamp, returns `current`.
 */
export function stampProperties(
  current: DocumentProperties | undefined,
  stamp: PropertiesStamp,
): DocumentProperties | undefined {
  if (!stamp.edited && !stamp.fresh) return current

  const when = w3cdtf(stamp.now)
  const author = stamp.author.trim()

  return {
    ...current,
    ...(stamp.fresh ? creationStamp(current, when, author) : {}),
    modified: when,
    ...(author === '' ? {} : { lastModifiedBy: author }),
    revision: nextRevision(current?.revision),
  }
}

/** A new document gets the creation date and author, without erasing existing ones. */
function creationStamp(
  current: DocumentProperties | undefined,
  when: string,
  author: string,
): Partial<DocumentProperties> {
  return {
    ...(current?.created === undefined ? { created: when } : {}),
    ...(author !== '' && (current?.creator ?? '') === '' ? { creator: author } : {}),
  }
}

function nextRevision(revision: string | undefined): string {
  const current = Number.parseInt(revision ?? '', 10)
  return String(Number.isFinite(current) && current > 0 ? current + 1 : 1)
}
