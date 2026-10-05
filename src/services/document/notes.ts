import { formatNumber } from './list-numbering.js'
import type { DocumentNotes, NoteNumbering } from './model.js'

/**
 * The number does not live in the reference: it is its order in the document, counted with the
 * format and start of `w:footnotePr`/`w:endnotePr`. A reference with its own mark
 * (`w:customMarkFollows`) is not counted, as in Word.
 */

export const NoteKind = {
  Footnote: 'footnote',
  Endnote: 'endnote',
} as const
export type NoteKind = (typeof NoteKind)[keyof typeof NoteKind]

export interface NoteReference {
  readonly kind: string
  readonly mark?: string | null
  /** The reference's section index, for `numRestart` `eachSect`. */
  readonly section?: number
  /** For `eachPage`. Without pagination, as in reading mode, the count runs on. */
  readonly page?: number
}

/** Word's default: 1, 2, 3 for footnotes; i, ii, iii for endnotes. */
const DEFAULT_FORMAT: Record<NoteKind, string> = {
  [NoteKind.Footnote]: 'decimal',
  [NoteKind.Endnote]: 'lowerRoman',
}

/** `chicago`: *, †, ‡, §, then the same doubled. */
const CHICAGO = ['*', '†', '‡', '§']

function numberingOf(kind: string, notes: DocumentNotes | undefined): NoteNumbering | undefined {
  return kind === NoteKind.Endnote ? notes?.endnotePr : notes?.footnotePr
}

/** `ordinal` from 0; restarting is `noteCounter`'s job. */
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
  // `none` and `bullet` cannot number: decimal is used.
  return text === '' ? String(value) : text
}

/**
 * Called in text order. `numRestart` goes back to `numStart` at the first note of each section or
 * each sheet (footnotes only, as in Word).
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
