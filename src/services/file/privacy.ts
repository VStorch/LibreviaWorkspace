import type { DocumentProperties } from './properties.js'

/** What the personal data inspector can remove. The sidecar uses the same names. */
export const PrivacyCategory = {
  /** Authorship in the properties, and the author of each comment and tracked change. */
  Names: 'names',
  /** Folders on the author's computer: the attached Word template, Excel's last folder. */
  Paths: 'paths',
  CustomProperties: 'customProperties',
  /** The page preview file managers show. */
  Thumbnail: 'thumbnail',
  /** EXIF (with the GPS position), XMP and comments inside JPEG and PNG images. */
  ImageMetadata: 'imageMetadata',
} as const

export type PrivacyCategory = (typeof PrivacyCategory)[keyof typeof PrivacyCategory]

export const PRIVACY_CATEGORIES = [
  PrivacyCategory.Names,
  PrivacyCategory.Paths,
  PrivacyCategory.CustomProperties,
  PrivacyCategory.Thumbnail,
  PrivacyCategory.ImageMetadata,
] as const

export interface PrivacyReport {
  readonly names: readonly string[]
  readonly paths: readonly string[]
  readonly customProperties: readonly string[]
  readonly thumbnail: boolean
  readonly imagesWithMetadata: number
  readonly removePersonalInformation: boolean
  /** A PDF's earlier revisions, which incremental updates leave in the file. */
  readonly previousVersions?: number
}

/** The categories with something to remove, in the order the inspector lists them. */
export function foundCategories(report: PrivacyReport): PrivacyCategory[] {
  const found: Record<PrivacyCategory, boolean> = {
    [PrivacyCategory.Names]: report.names.length > 0,
    [PrivacyCategory.Paths]: report.paths.length > 0,
    [PrivacyCategory.CustomProperties]: report.customProperties.length > 0,
    [PrivacyCategory.Thumbnail]: report.thumbnail,
    [PrivacyCategory.ImageMetadata]: report.imagesWithMetadata > 0,
  }
  return PRIVACY_CATEGORIES.filter((category) => found[category])
}

/** The properties that name people or the company, as Word's inspector removes them. */
const PERSONAL_PROPERTIES = ['creator', 'lastModifiedBy', 'company', 'manager'] as const

/**
 * A `.sdoc` or `.ssheet` only has names: in the properties, and in the comments and changes of
 * `content` (the document model, or nothing for a spreadsheet).
 */
export function inspectOwnFile(
  properties: DocumentProperties | undefined,
  content: unknown,
  anonymous: string,
): PrivacyReport {
  const fromProperties = PERSONAL_PROPERTIES.flatMap((key) => (properties?.[key] ?? '').split(';'))
  const names = [...new Set([...fromProperties, ...authorsIn(content)].map((name) => name.trim()))].filter(
    (name) => name !== '' && name !== anonymous,
  )

  return {
    names,
    paths: [],
    customProperties: [],
    thumbnail: false,
    imagesWithMetadata: 0,
    removePersonalInformation: properties?.removePersonalInformation === true,
  }
}

/**
 * A PDF only shows its author here: changing it appends an update and leaves the old value in the
 * file, so the inspector warns instead of promising a removal.
 */
export function inspectPdf(
  properties: DocumentProperties | undefined,
  updates: number,
  anonymous: string,
): PrivacyReport {
  return { ...inspectOwnFile(properties, null, anonymous), previousVersions: updates }
}

/** Empty clears the field when saving to DOCX or XLSX; the flag keeps the next save from signing. */
export function withoutPersonalNames(properties: DocumentProperties | undefined): DocumentProperties {
  return {
    ...properties,
    ...Object.fromEntries(PERSONAL_PROPERTIES.map((key) => [key, ''])),
    removePersonalInformation: true,
  }
}

const AUTHOR_KEY = 'author'
const INITIALS_KEY = 'initials'

/** Every `author` in a JSON tree: comments, revision marks, paragraph mark revisions. */
export function authorsIn(value: unknown): string[] {
  if (Array.isArray(value)) return value.flatMap(authorsIn)
  if (typeof value !== 'object' || value === null) return []

  return Object.entries(value).flatMap(([key, child]) =>
    key === AUTHOR_KEY && typeof child === 'string' ? [child] : authorsIn(child),
  )
}

/** The same tree with every author, and the comments' initials, made generic, as Word does. */
export function anonymizeAuthors<T>(value: T, anonymous: string): T {
  return anonymizeNode(value, anonymous) as T
}

function anonymizeNode(value: unknown, anonymous: string): unknown {
  if (Array.isArray(value)) return value.map((child) => anonymizeNode(child, anonymous))
  if (typeof value !== 'object' || value === null) return value

  return Object.fromEntries(
    Object.entries(value).map(([key, child]) => {
      if (key === AUTHOR_KEY && typeof child === 'string') return [key, anonymous]
      if (key === INITIALS_KEY && typeof child === 'string') return [key, anonymous.slice(0, 1)]
      return [key, anonymizeNode(child, anonymous)]
    }),
  )
}
