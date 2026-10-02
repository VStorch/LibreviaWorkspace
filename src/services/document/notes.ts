import { formatNumber } from './list-numbering.js'
import type { DocumentNotes, NoteNumbering } from './model.js'

/**
 * A numeração das notas de rodapé e de fim (M11).
 *
 * O número não mora na referência (`noteRef`): é a ordem dela no documento, e
 * guardado ele envelheceria na primeira nota inserida antes. Quem o desenha — a
 * tela e o papel — conta aqui, com o formato e o início que o documento declara
 * (`w:footnotePr`/`w:endnotePr`, em `DocumentModel.notes`).
 *
 * A referência de marca própria (`w:customMarkFollows`) mostra a marca e não
 * entra na conta, como no Word.
 */

export const NoteKind = {
  Footnote: 'footnote',
  Endnote: 'endnote',
} as const
export type NoteKind = (typeof NoteKind)[keyof typeof NoteKind]

/** O que a conta precisa de cada referência, na ordem do documento. */
export interface NoteReference {
  readonly kind: string
  readonly mark?: string | null
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

/**
 * O rótulo da nota numerada de índice `ordinal` (a partir de 0) do tipo dado.
 *
 * `numRestart` (por seção, por página) ainda não muda a conta: sem paginar as
 * notas, a tela não sabe em que página cada uma cai — é a fase 2.
 */
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
  // `none` e `bullet` não servem de número de nota: o número certo em decimal é
  // melhor que a referência sumir.
  return text === '' ? String(value) : text
}

/** O rótulo de cada referência, na mesma ordem. */
export function noteLabels(references: readonly NoteReference[], notes?: DocumentNotes): string[] {
  const counters = new Map<string, number>()
  return references.map((reference) => {
    if (typeof reference.mark === 'string' && reference.mark !== '') return reference.mark
    const ordinal = counters.get(reference.kind) ?? 0
    counters.set(reference.kind, ordinal + 1)
    return noteLabel(reference.kind, ordinal, notes)
  })
}
