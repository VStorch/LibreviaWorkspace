import type { z } from 'zod'
import type { documentPropertiesSchema } from '@shared/schemas.js'

export const CustomPropertyType = {
  Text: 'text',
  Number: 'number',
  /** The value is the file's instant (`2026-10-12T03:00:00Z`): Word stores local midnight in UTC. */
  Date: 'date',
  /** `true` or `false`. */
  Boolean: 'boolean',
} as const

export type CustomPropertyType = (typeof CustomPropertyType)[keyof typeof CustomPropertyType]

export const CUSTOM_PROPERTY_TYPES = [
  CustomPropertyType.Text,
  CustomPropertyType.Number,
  CustomPropertyType.Date,
  CustomPropertyType.Boolean,
] as const

/** File → Properties → Custom, as `docProps/custom.xml` keeps it. */
export interface CustomProperty {
  readonly name: string
  readonly type: CustomPropertyType
  readonly value: string
}

/**
 * As Word and Excel show them in File → Properties. Dates are W3CDTF (`2026-10-02T12:00:00Z`), as in the
 * package.
 */
export interface DocumentProperties {
  readonly title?: string
  readonly subject?: string
  /** `dc:creator`: the author(s), separated by semicolons, as in Word. */
  readonly creator?: string
  readonly keywords?: string
  readonly category?: string
  /** `cp:contentStatus`: "Rascunho", "Final"… */
  readonly contentStatus?: string
  /** `dc:language`, as Word wrote it (`pt-BR`). Read only. */
  readonly language?: string
  /** `cp:lastPrinted`. Read only: the editor does not mark printing. */
  readonly lastPrinted?: string
  /** `dc:description`: what Word calls Comments. */
  readonly description?: string
  readonly lastModifiedBy?: string
  readonly revision?: string
  readonly created?: string
  readonly modified?: string
  /** `docProps/app.xml`. */
  readonly company?: string
  readonly manager?: string
  /** `TotalTime` from `app.xml`, in minutes. Read only: the editor does not measure it. */
  readonly totalTime?: number
  /**
   * Word's `w:removePersonalInformation`, Excel's `filterPrivacy`: saving signs nothing with the
   * person's name. Set by the personal data inspector.
   */
  readonly removePersonalInformation?: boolean
  /** Absent leaves the file's as they are; a list replaces them. */
  readonly custom?: readonly CustomProperty[]
}

/**
 * What Word keeps in `app.xml`, for file managers and search that do not open the document. Only
 * sent when saving an edited DOCX: never read back, nor kept with the properties.
 */
export interface DocumentStatistics {
  readonly pages: number
  readonly words: number
  /** Without spaces, as Word's `Characters`. */
  readonly characters: number
  readonly charactersWithSpaces: number
  readonly paragraphs: number
}

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
  const author = current?.removePersonalInformation === true ? '' : stamp.author.trim()

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

/** Without absent keys, for `exactOptionalPropertyTypes`. */
export function propertiesOf(raw: z.infer<typeof documentPropertiesSchema>): DocumentProperties {
  return Object.fromEntries(
    Object.entries(raw).filter(([, value]) => value !== undefined),
  ) as DocumentProperties
}

export type CustomPropertyProblem = 'nameMissing' | 'nameRepeated' | 'numberInvalid' | 'dateMissing'

/** The first problem that would keep the list from being saved; `null` when there is none. */
export function customPropertiesProblem(properties: readonly CustomProperty[]): CustomPropertyProblem | null {
  const names = properties.map((property) => property.name.trim().toLocaleLowerCase())
  if (names.some((name) => name === '')) return 'nameMissing'
  // Word refuses two properties that differ only in case.
  if (new Set(names).size !== names.length) return 'nameRepeated'
  if (
    properties.some((property) => property.type === CustomPropertyType.Number && !isNumber(property.value))
  ) {
    return 'numberInvalid'
  }
  if (properties.some((property) => property.type === CustomPropertyType.Date && property.value === '')) {
    return 'dateMissing'
  }
  return null
}

/** What a number property accepts: Word refuses an empty value. */
export function isNumber(value: string): boolean {
  return value.trim() !== '' && Number.isFinite(Number(value))
}

/** In any order: the file's list and the one on screen need not follow each other. */
export function sameCustomProperties(a: readonly CustomProperty[], b: readonly CustomProperty[]): boolean {
  return (
    a.length === b.length &&
    a.every((left) =>
      b.some((right) => right.name === left.name && right.type === left.type && right.value === left.value),
    )
  )
}

/** The day the person sees, in their time zone, as a date input wants it (`2026-10-12`). */
export function localDayOf(instant: string): string {
  const parsed = Date.parse(instant)
  if (Number.isNaN(parsed)) return ''
  const date = new Date(parsed)
  const pad = (value: number): string => String(value).padStart(2, '0')
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
}

/** Local midnight of the day, in UTC, as Word writes a custom date. */
export function instantOfDay(day: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(day)
  if (match === null) return ''
  const [, year, month, date] = match
  return w3cdtf(new Date(Number(year), Number(month) - 1, Number(date)))
}
