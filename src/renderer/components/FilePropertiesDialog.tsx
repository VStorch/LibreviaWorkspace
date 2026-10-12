import { useState } from 'react'
import {
  customPropertiesProblem,
  sameCustomProperties,
  type CustomProperty,
  type CustomPropertyProblem,
  type DocumentProperties,
} from '@services/file/properties.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { MAX_DESCRIPTION_LENGTH, MAX_PROPERTY_LENGTH } from '@shared/limits.js'
import { useLanguage, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { CustomPropertiesFields, propertiesOfRows, rowsOf, type CustomRow } from './CustomPropertiesFields.js'
import { DialogActions } from './DialogActions.js'

/** A label and the value already written for the reader. */
export type StatisticRow = readonly [MessageKey, string]

/** In Word's order. */
const EDITABLE = [
  ['title', 'document.properties.title'],
  ['subject', 'document.properties.subject'],
  ['creator', 'document.properties.creator'],
  ['manager', 'document.properties.manager'],
  ['company', 'document.properties.company'],
  ['category', 'document.properties.category'],
  ['keywords', 'document.properties.keywords'],
  ['description', 'document.properties.description'],
] as const satisfies readonly (readonly [keyof DocumentProperties, MessageKey])[]

export type EditableKey = (typeof EDITABLE)[number][0]

const ALL_FIELDS: readonly EditableKey[] = EDITABLE.map(([key]) => key)

/**
 * File → Properties, for a document or a spreadsheet: the same fields, each editor with its own
 * statistics. Saves a **patch**: a field that did not exist and stayed empty stays absent; one that
 * was cleared becomes an empty string. Dates, last saved by and revision are read only
 * (`stampProperties`).
 */
export function FilePropertiesDialog({
  title,
  statistics,
  onClose,
  fields = ALL_FIELDS,
  authorship = true,
  custom = true,
  notice,
}: {
  readonly title: MessageKey
  readonly statistics: readonly StatisticRow[]
  readonly onClose: () => void
  /** The format's own: a PDF has no company, manager, category or comments. */
  readonly fields?: readonly EditableKey[]
  /** Last saved by, revision and editing time, which only Office files keep. */
  readonly authorship?: boolean
  /** `docProps/custom.xml`, which a PDF does not have. */
  readonly custom?: boolean
  readonly notice?: string | undefined
}): React.JSX.Element {
  const t = useT()
  const properties = useWorkspace((state) => state.properties)
  const setProperties = useWorkspace((state) => state.setProperties)
  const [values, setValues] = useState<Record<EditableKey, string>>(
    () =>
      Object.fromEntries(EDITABLE.map(([key]) => [key, properties?.[key] ?? ''])) as Record<
        EditableKey,
        string
      >,
  )

  const [rows, setRows] = useState<CustomRow[]>(() => rowsOf(properties?.custom))
  const [problem, setProblem] = useState<CustomPropertyProblem | null>(null)

  function save(): void {
    const list = custom ? propertiesOfRows(rows) : null
    const found = list === null ? null : customPropertiesProblem(list)
    setProblem(found)
    if (found !== null) return

    const next = patchOf(properties, values, list)
    if (next !== null) setProperties(next)
    onClose()
  }

  return (
    <div
      className="popover popover--wide properties"
      role="dialog"
      aria-label={t(title)}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      {notice !== undefined && (
        <p className="popover__hint" role="note">
          {notice}
        </p>
      )}

      <SummaryFields fields={fields} values={values} onChange={setValues} onSubmit={save} />

      {custom && <CustomPropertiesFields rows={rows} onChange={setRows} problem={problem} />}

      <Statistics statistics={statistics} authorship={authorship} />

      <DialogActions confirmLabel={t('document.common.apply')} onConfirm={save} onCancel={onClose} />
    </div>
  )
}

/** `null` when nothing changed. `custom` is `null` where the format has none. */
function patchOf(
  properties: DocumentProperties | undefined,
  values: Record<EditableKey, string>,
  custom: readonly CustomProperty[] | null,
): DocumentProperties | null {
  let next: Record<string, unknown> = { ...properties }
  let changed = custom !== null && !sameCustomProperties(properties?.custom ?? [], custom)
  if (changed) next = { ...next, custom }
  for (const [key] of EDITABLE) {
    const value = values[key]
    if (value === (properties?.[key] ?? '')) continue
    next = { ...next, [key]: value }
    changed = true
  }
  return changed ? (next as DocumentProperties) : null
}

function SummaryFields({
  fields,
  values,
  onChange,
  onSubmit,
}: {
  fields: readonly EditableKey[]
  values: Record<EditableKey, string>
  onChange: (values: Record<EditableKey, string>) => void
  onSubmit: () => void
}): React.JSX.Element {
  const t = useT()
  return (
    <fieldset className="popover__fieldset">
      <legend>{t('document.properties.summary')}</legend>
      {EDITABLE.filter(([key]) => fields.includes(key)).map(([key, label], index) => (
        <label key={key} className="popover__field properties__field">
          <span>{t(label)}</span>
          {key === 'description' ? (
            <textarea
              rows={3}
              value={values[key]}
              maxLength={MAX_DESCRIPTION_LENGTH}
              onChange={(event) => onChange({ ...values, [key]: event.target.value })}
            />
          ) : (
            <input
              type="text"
              value={values[key]}
              maxLength={MAX_PROPERTY_LENGTH}
              autoFocus={index === 0}
              onChange={(event) => onChange({ ...values, [key]: event.target.value })}
              onKeyDown={(event) => {
                if (event.key === 'Enter') onSubmit()
              }}
            />
          )}
        </label>
      ))}
    </fieldset>
  )
}

function Statistics({
  statistics,
  authorship,
}: {
  statistics: readonly StatisticRow[]
  authorship: boolean
}): React.JSX.Element {
  const t = useT()
  const language = useLanguage()
  const properties = useWorkspace((state) => state.properties)

  const date = (value: string | undefined): string => {
    const parsed = value === undefined || value === '' ? Number.NaN : Date.parse(value)
    return Number.isNaN(parsed)
      ? '—'
      : new Date(parsed).toLocaleString(language, { dateStyle: 'short', timeStyle: 'short' })
  }

  const info: readonly StatisticRow[] = [
    ['document.properties.created', date(properties?.created)],
    ['document.properties.modified', date(properties?.modified)],
    ...(authorship ? authorshipRows(properties, t, language) : []),
    ...statistics,
  ]

  return (
    <fieldset className="popover__fieldset">
      <legend>{t('document.properties.statistics')}</legend>
      <table className="counts">
        <tbody>
          {info.map(([label, value]) => (
            <tr key={label}>
              <th scope="row">{t(label)}</th>
              <td>{value}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </fieldset>
  )
}

function authorshipRows(
  properties: DocumentProperties | undefined,
  t: ReturnType<typeof useT>,
  language: string,
): StatisticRow[] {
  return [
    ['document.properties.lastModifiedBy', properties?.lastModifiedBy || '—'],
    ['document.properties.revision', properties?.revision || '—'],
    ...(properties?.totalTime === undefined
      ? []
      : [
          [
            'document.properties.totalTime',
            t('document.properties.minutes', { count: properties.totalTime.toLocaleString(language) }),
          ] as const,
        ]),
  ]
}
