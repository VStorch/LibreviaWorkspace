import {
  CUSTOM_PROPERTY_TYPES,
  CustomPropertyType,
  instantOfDay,
  localDayOf,
  type CustomProperty,
  type CustomPropertyProblem,
} from '@services/file/properties.js'
import { MAX_NAME_LENGTH, MAX_PROPERTY_LENGTH } from '@shared/limits.js'
import { useT } from '../i18n.js'

/** A property on screen; `key` only keeps React's rows apart while names are being typed. */
export interface CustomRow extends CustomProperty {
  readonly key: number
}

export function rowsOf(properties: readonly CustomProperty[] | undefined): CustomRow[] {
  return (properties ?? []).map((property, index) => ({ ...property, key: index }))
}

/** The list without the screen's keys, with names as they will be saved. */
export function propertiesOfRows(rows: readonly CustomRow[]): CustomProperty[] {
  return rows.map(({ name, type, value }) => ({ name: name.trim(), type, value }))
}

/** What a value becomes when its type changes: what still fits stays. */
function valueForType(type: CustomPropertyType, previous: string): string {
  switch (type) {
    case CustomPropertyType.Boolean:
      return 'false'
    case CustomPropertyType.Date:
      return ''
    case CustomPropertyType.Number:
      return previous.trim() !== '' && Number.isFinite(Number(previous)) ? previous : ''
    case CustomPropertyType.Text:
      return previous
  }
}

/** File → Properties → Custom: name, type and value, as in Word. */
export function CustomPropertiesFields({
  rows,
  onChange,
  problem,
}: {
  readonly rows: readonly CustomRow[]
  readonly onChange: (rows: CustomRow[]) => void
  readonly problem: CustomPropertyProblem | null
}): React.JSX.Element {
  const t = useT()
  const update = (key: number, change: Partial<CustomProperty>): void =>
    onChange(rows.map((row) => (row.key === key ? { ...row, ...change } : row)))
  const add = (): void => {
    const key = rows.reduce((highest, row) => Math.max(highest, row.key), -1) + 1
    onChange([...rows, { key, name: '', type: CustomPropertyType.Text, value: '' }])
  }

  return (
    <fieldset className="popover__fieldset">
      <legend>{t('document.properties.custom')}</legend>
      {rows.length === 0 ? (
        <p className="popover__hint">{t('document.properties.custom.empty')}</p>
      ) : (
        <div className="custom-properties">
          <span className="custom-properties__header">{t('document.properties.custom.name')}</span>
          <span className="custom-properties__header">{t('document.properties.custom.type')}</span>
          <span className="custom-properties__header">{t('document.properties.custom.value')}</span>
          <span />
          {rows.map((row, index) => (
            <CustomRowFields
              key={row.key}
              row={row}
              focus={index === rows.length - 1 && row.name === ''}
              onChange={(change) => update(row.key, change)}
              onRemove={() => onChange(rows.filter((other) => other.key !== row.key))}
            />
          ))}
        </div>
      )}
      {problem !== null && (
        <p className="popover__error" role="alert">
          {t(`document.properties.custom.${problem}`)}
        </p>
      )}
      <button type="button" className="btn" onClick={add}>
        {t('document.properties.custom.add')}
      </button>
    </fieldset>
  )
}

function CustomRowFields({
  row,
  focus,
  onChange,
  onRemove,
}: {
  readonly row: CustomRow
  readonly focus: boolean
  readonly onChange: (change: Partial<CustomProperty>) => void
  readonly onRemove: () => void
}): React.JSX.Element {
  const t = useT()
  return (
    <div className="custom-properties__row">
      <input
        type="text"
        aria-label={t('document.properties.custom.name')}
        value={row.name}
        maxLength={MAX_NAME_LENGTH}
        autoFocus={focus}
        onChange={(event) => onChange({ name: event.target.value })}
      />
      <select
        aria-label={t('document.properties.custom.type')}
        value={row.type}
        onChange={(event) => {
          const type = event.target.value as CustomPropertyType
          onChange({ type, value: valueForType(type, row.value) })
        }}
      >
        {CUSTOM_PROPERTY_TYPES.map((type) => (
          <option key={type} value={type}>
            {t(`document.properties.custom.${type}`)}
          </option>
        ))}
      </select>
      <ValueField row={row} onChange={(value) => onChange({ value })} />
      <button
        type="button"
        className="btn"
        aria-label={t('document.properties.custom.remove', { name: row.name })}
        onClick={onRemove}
      >
        ×
      </button>
    </div>
  )
}

function ValueField({
  row,
  onChange,
}: {
  readonly row: CustomRow
  readonly onChange: (value: string) => void
}): React.JSX.Element {
  const t = useT()
  const label = t('document.properties.custom.value')
  switch (row.type) {
    case CustomPropertyType.Boolean:
      return (
        <input
          type="checkbox"
          aria-label={label}
          checked={row.value === 'true'}
          onChange={(event) => onChange(event.target.checked ? 'true' : 'false')}
        />
      )
    case CustomPropertyType.Date:
      return (
        <input
          type="date"
          aria-label={label}
          value={localDayOf(row.value)}
          onChange={(event) => onChange(instantOfDay(event.target.value))}
        />
      )
    case CustomPropertyType.Number:
      return (
        <input
          type="number"
          step="any"
          aria-label={label}
          value={row.value}
          onChange={(event) => onChange(event.target.value)}
        />
      )
    case CustomPropertyType.Text:
      return (
        <input
          type="text"
          aria-label={label}
          value={row.value}
          maxLength={MAX_PROPERTY_LENGTH}
          onChange={(event) => onChange(event.target.value)}
        />
      )
  }
}
