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
  /** A seção da referência (o índice dela), para `numRestart` `eachSect`. */
  readonly section?: number
  /**
   * A folha da referência, para `eachPage`. Só a paginação a sabe: antes dela (e
   * no modo de leitura, sem folha) a conta segue, como se fosse contínua.
   */
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

/**
 * O rótulo da nota numerada de índice `ordinal` (a partir de 0) do tipo dado.
 *
 * O reinício (`numRestart`) é de quem conta — ver `noteCounter`.
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

/**
 * A conta que corre pelo documento: devolve o rótulo de cada referência, chamada
 * na ordem do texto. `numRestart` volta ao início (`numStart`) na primeira nota
 * de cada seção (`eachSect`) ou de cada folha (`eachPage`, só nas de rodapé —
 * a de fim não tem folha própria, e o Word não oferece). Referência sem a seção
 * ou a folha conhecida não reinicia.
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

/** O rótulo de cada referência, na mesma ordem. */
export function noteLabels(references: readonly NoteReference[], notes?: DocumentNotes): string[] {
  return references.map(noteCounter(notes))
}
