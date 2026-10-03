import { formatNumber } from './list-numbering.js'
import type { DocumentNotes, NoteNumbering } from './model.js'

/**
 * O número não mora na referência: é a ordem dela no documento, contada com o
 * formato e o início de `w:footnotePr`/`w:endnotePr`. A referência de marca
 * própria (`w:customMarkFollows`) não entra na conta, como no Word.
 */

export const NoteKind = {
  Footnote: 'footnote',
  Endnote: 'endnote',
} as const
export type NoteKind = (typeof NoteKind)[keyof typeof NoteKind]

export interface NoteReference {
  readonly kind: string
  readonly mark?: string | null
  /** A seção da referência (o índice dela), para `numRestart` `eachSect`. */
  readonly section?: number
  /** Para `eachPage`. Sem paginação, como no modo de leitura, a conta segue contínua. */
  readonly page?: number
}

/** O padrão do Word: 1, 2, 3 nas de rodapé; i, ii, iii nas de fim. */
const DEFAULT_FORMAT: Record<NoteKind, string> = {
  [NoteKind.Footnote]: 'decimal',
  [NoteKind.Endnote]: 'lowerRoman',
}

/** `chicago`: *, †, ‡, § — e depois as mesmas, dobradas. */
const CHICAGO = ['*', '†', '‡', '§']

function numberingOf(kind: string, notes: DocumentNotes | undefined): NoteNumbering | undefined {
  return kind === NoteKind.Endnote ? notes?.endnotePr : notes?.footnotePr
}

/** `ordinal` a partir de 0; o reinício é de `noteCounter`. */
export function noteLabel(kind: string, ordinal: number, notes?: DocumentNotes): string {
  const numbering = numberingOf(kind, notes)
  const value = (numbering?.start ?? 1) + ordinal
  const format =
    numbering?.numFmt ?? DEFAULT_FORMAT[kind === NoteKind.Endnote ? NoteKind.Endnote : NoteKind.Footnote]
  if (format === 'chicago') {
    if (value <= 0) return String(value)
    return CHICAGO[(value - 1) % CHICAGO.length]!.repeat(Math.floor((value - 1) / CHICAGO.length) + 1)
  }
  const text = formatNumber(value, format)
  // `none` e `bullet` não servem de número: sai em decimal.
  return text === '' ? String(value) : text
}

/**
 * Chamada na ordem do texto. `numRestart` volta ao `numStart` na primeira nota
 * de cada seção ou de cada folha (só nas de rodapé, como no Word).
 */
export function noteCounter(notes?: DocumentNotes): (reference: NoteReference) => string {
  const last = new Map<string, { ordinal: number; section?: number; page?: number }>()
  return (reference) => {
    if (typeof reference.mark === 'string' && reference.mark !== '') return reference.mark
    const restart = numberingOf(reference.kind, notes)?.restart
    const previous = last.get(reference.kind)
    const moved = (before: number | undefined, now: number | undefined): boolean =>
      before !== undefined && now !== undefined && before !== now
    const ordinal =
      previous === undefined ||
      (restart === 'eachSect' && moved(previous.section, reference.section)) ||
      (restart === 'eachPage' && reference.kind !== NoteKind.Endnote && moved(previous.page, reference.page))
        ? 0
        : previous.ordinal + 1
    last.set(reference.kind, {
      ordinal,
      ...(reference.section === undefined ? {} : { section: reference.section }),
      ...(reference.page === undefined ? {} : { page: reference.page }),
    })
    return noteLabel(reference.kind, ordinal, notes)
  }
}

export function noteLabels(references: readonly NoteReference[], notes?: DocumentNotes): string[] {
  return references.map(noteCounter(notes))
}
