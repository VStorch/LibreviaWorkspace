import { CustomPropertyType, type DocumentProperties } from '@services/file/properties.js'
import { fieldKind, fieldTokens } from './fields.js'

/**
 * Fields that show a property of the file: `TITLE`, `AUTHOR`… and `DOCPROPERTY "Cliente"`, which
 * also reaches the built-in ones by Word's English name.
 */
export const PROPERTY_FIELD_KINDS: ReadonlySet<string> = new Set([
  'TITLE',
  'SUBJECT',
  'AUTHOR',
  'KEYWORDS',
  'COMMENTS',
  'LASTSAVEDBY',
  'DOCPROPERTY',
])

export type BuiltInKey =
  | 'title'
  | 'subject'
  | 'creator'
  | 'keywords'
  | 'description'
  | 'category'
  | 'company'
  | 'manager'
  | 'lastModifiedBy'

/** Word's names, lowercase; the dedicated fields have the same ones. */
const BUILT_IN: Readonly<Record<string, BuiltInKey>> = {
  title: 'title',
  subject: 'subject',
  author: 'creator',
  keywords: 'keywords',
  comments: 'description',
  category: 'category',
  company: 'company',
  manager: 'manager',
  lastsavedby: 'lastModifiedBy',
}

/** Word writes a yes-or-no property this way in any language. */
const YES = 'Y'
const NO = 'N'

/**
 * What the field shows now; `null` for a `DOCPROPERTY` naming a property the file does not have,
 * which Word shows as an error. `locale` formats dates and numbers, as Word follows the system.
 */
export function propertyFieldResult(
  instr: string,
  properties: DocumentProperties | undefined,
  locale: string,
): string | null {
  const kind = fieldKind(instr)
  const name = (kind === 'DOCPROPERTY' ? (fieldTokens(instr)[1] ?? '') : kind).toLowerCase()
  const key = BUILT_IN[name]
  if (key !== undefined) return withCase(properties?.[key] ?? '', instr)

  const custom = properties?.custom?.find((property) => property.name.toLowerCase() === name)
  if (custom === undefined) return null
  return withCase(customText(custom.type, custom.value, locale), instr)
}

function customText(type: CustomPropertyType, value: string, locale: string): string {
  switch (type) {
    case CustomPropertyType.Boolean:
      return value === 'true' ? YES : NO
    case CustomPropertyType.Date: {
      const parsed = Date.parse(value)
      return Number.isNaN(parsed) ? value : new Date(parsed).toLocaleDateString(locale)
    }
    case CustomPropertyType.Number: {
      const number = Number(value)
      return Number.isFinite(number)
        ? number.toLocaleString(locale, { useGrouping: false, maximumFractionDigits: 20 })
        : value
    }
    case CustomPropertyType.Text:
      return value
  }
}

/** `\* Upper`, `\* Lower`, `\* Caps` and `\* FirstCap`; `\* MERGEFORMAT` only keeps formatting. */
function withCase(text: string, instr: string): string {
  const tokens = fieldTokens(instr)
  const formats = tokens.flatMap((token, index) => (token === '\\*' ? [tokens[index + 1] ?? ''] : []))
  return formats.reduce((current, format) => {
    switch (format.toLowerCase()) {
      case 'upper':
        return current.toLocaleUpperCase()
      case 'lower':
        return current.toLocaleLowerCase()
      case 'caps':
        return current.replace(
          /(^|\s)(\S)/gu,
          (_, space: string, letter: string) => space + letter.toLocaleUpperCase(),
        )
      case 'firstcap':
        return current.charAt(0).toLocaleUpperCase() + current.slice(1)
      default:
        return current
    }
  }, text)
}

/** What "Insert → Document property" offers: the summary fields, then the custom ones. */
export type PropertyFieldChoice =
  | { readonly builtIn: true; readonly key: BuiltInKey; readonly instr: string }
  | { readonly builtIn: false; readonly name: string; readonly instr: string }

/** Word's own field where there is one, `DOCPROPERTY` for the rest. */
const BUILT_IN_CHOICES: readonly { readonly key: BuiltInKey; readonly instr: string }[] = [
  { key: 'title', instr: 'TITLE' },
  { key: 'subject', instr: 'SUBJECT' },
  { key: 'creator', instr: 'AUTHOR' },
  { key: 'keywords', instr: 'KEYWORDS' },
  { key: 'description', instr: 'COMMENTS' },
  { key: 'category', instr: 'DOCPROPERTY Category' },
  { key: 'company', instr: 'DOCPROPERTY Company' },
  { key: 'manager', instr: 'DOCPROPERTY Manager' },
]

/** Word inserts every field with `\* MERGEFORMAT`: the result keeps the formatting given to it. */
const MERGEFORMAT = '\\* MERGEFORMAT'

export function propertyFieldChoices(properties: DocumentProperties | undefined): PropertyFieldChoice[] {
  return [
    ...BUILT_IN_CHOICES.map(({ key, instr }) => ({
      builtIn: true as const,
      key,
      instr: ` ${instr} ${MERGEFORMAT} `,
    })),
    // A quote would end the name inside the instruction.
    ...(properties?.custom ?? [])
      .filter((property) => !property.name.includes('"'))
      .map((property) => ({
        builtIn: false as const,
        name: property.name,
        instr: ` DOCPROPERTY "${property.name}" ${MERGEFORMAT} `,
      })),
  ]
}
