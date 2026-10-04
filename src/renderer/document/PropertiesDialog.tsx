import { useState } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import type { DocumentProperties } from '@services/document/model.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { useLanguage, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { DialogActions } from '../components/DialogActions.js'
import { MAX_DESCRIPTION_LENGTH, MAX_PROPERTY_LENGTH } from '@shared/limits.js'
import { tally } from './WordCountDialog.js'

/** Na ordem do Word. */
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

type EditableKey = (typeof EDITABLE)[number][0]

/**
 * Grava um **remendo**: o campo que não existia e ficou vazio fica ausente; o que
 * foi apagado vira cadeia vazia. Datas, autor da última gravação e revisão são só
 * lidos (`stampProperties`).
 */
export function PropertiesDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
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

  function save(): void {
    const next = patchOf(properties, values)
    if (next !== null) setProperties(next)
    close()
  }

  function close(): void {
    onClose()
    requestAnimationFrame(() => editor.view.focus())
  }

  return (
    <div
      className="popover popover--wide properties"
      role="dialog"
      aria-label={t('document.properties.title.dialog')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') close()
      }}
    >
      <SummaryFields values={values} onChange={setValues} onSubmit={save} />

      <Statistics editor={editor} />

      <DialogActions confirmLabel={t('document.common.apply')} onConfirm={save} onCancel={close} />
    </div>
  )
}

/** `null` quando nada mudou. */
function patchOf(
  properties: DocumentProperties | undefined,
  values: Record<EditableKey, string>,
): DocumentProperties | null {
  let next: Record<string, unknown> = { ...properties }
  let changed = false
  for (const [key] of EDITABLE) {
    const value = values[key]
    if (value === (properties?.[key] ?? '')) continue
    next = { ...next, [key]: value }
    changed = true
  }
  return changed ? (next as DocumentProperties) : null
}

function SummaryFields({
  values,
  onChange,
  onSubmit,
}: {
  values: Record<EditableKey, string>
  onChange: (values: Record<EditableKey, string>) => void
  onSubmit: () => void
}): React.JSX.Element {
  const t = useT()
  return (
    <fieldset className="popover__fieldset">
      <legend>{t('document.properties.summary')}</legend>
      {EDITABLE.map(([key, label], index) => (
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

function Statistics({ editor }: { editor: Editor }): React.JSX.Element {
  const t = useT()
  const language = useLanguage()
  const properties = useWorkspace((state) => state.properties)
  const pages = useWorkspace((state) => state.pageCount)
  const counts = useEditorState({
    editor,
    selector: ({ editor: current }) => tally(current, current.state.doc),
  })

  const date = (value: string | undefined): string => {
    const parsed = value === undefined || value === '' ? Number.NaN : Date.parse(value)
    return Number.isNaN(parsed)
      ? '—'
      : new Date(parsed).toLocaleString(language, { dateStyle: 'short', timeStyle: 'short' })
  }
  const number = (value: number): string => value.toLocaleString(language)

  const info: readonly (readonly [MessageKey, string])[] = [
    ['document.properties.created', date(properties?.created)],
    ['document.properties.modified', date(properties?.modified)],
    ['document.properties.lastModifiedBy', properties?.lastModifiedBy || '—'],
    ['document.properties.revision', properties?.revision || '—'],
    ...(properties?.totalTime === undefined
      ? []
      : [
          [
            'document.properties.totalTime',
            t('document.properties.minutes', { count: number(properties.totalTime) }),
          ] as const,
        ]),
    ['document.wordCount.pages', number(pages)],
    ['document.wordCount.words', number(counts.words)],
    ['document.wordCount.charactersWithSpaces', number(counts.characters)],
    ['document.wordCount.charactersNoSpaces', number(counts.charactersNoSpaces)],
    ['document.wordCount.paragraphs', number(counts.paragraphs)],
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
