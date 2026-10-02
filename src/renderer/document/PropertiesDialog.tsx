import { useState } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import type { DocumentProperties } from '@services/document/model.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { useLanguage, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { tally } from './WordCountDialog.js'

/** Os campos que o diálogo deixa editar, na ordem do Word. */
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
 * Arquivo → Propriedades (M11): o resumo que o Word guarda em `docProps/` e as
 * estatísticas do documento.
 *
 * O que se grava é um **remendo**: o campo que o arquivo não tinha e continuou
 * vazio fica ausente, e o que tinha e foi apagado vira cadeia vazia — "apague".
 * Assim o sidecar só regrava a parte de propriedades quando algum campo mudou
 * de fato.
 *
 * As estatísticas são as da contagem de palavras, pela mesma régua; as datas, o
 * autor da última gravação e a revisão são só lidos — quem os escreve é a
 * gravação (ver `stampProperties`).
 */
export function PropertiesDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const language = useLanguage()
  const properties = useWorkspace((state) => state.properties)
  const pages = useWorkspace((state) => state.estimatedPages)
  const setProperties = useWorkspace((state) => state.setProperties)
  const [values, setValues] = useState<Record<EditableKey, string>>(
    () =>
      Object.fromEntries(EDITABLE.map(([key]) => [key, properties?.[key] ?? ''])) as Record<
        EditableKey,
        string
      >,
  )

  const counts = useEditorState({
    editor,
    selector: ({ editor: current }) => tally(current, current.state.doc),
  })

  function save(): void {
    let next: Record<string, unknown> = { ...properties }
    let changed = false
    for (const [key] of EDITABLE) {
      const before = properties?.[key]
      const value = values[key]
      // O campo que o arquivo não tinha e continua vazio não vira "apague".
      if (value === (before ?? '')) continue
      next = { ...next, [key]: value }
      changed = true
    }
    if (changed) setProperties(next as DocumentProperties)
    close()
  }

  function close(): void {
    onClose()
    requestAnimationFrame(() => editor.view.focus())
  }

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
    <div
      className="popover popover--wide properties"
      role="dialog"
      aria-label={t('document.properties.title.dialog')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') close()
      }}
    >
      <fieldset className="popover__fieldset">
        <legend>{t('document.properties.summary')}</legend>
        {EDITABLE.map(([key, label], index) => (
          <label key={key} className="popover__field properties__field">
            <span>{t(label)}</span>
            {key === 'description' ? (
              <textarea
                rows={3}
                value={values[key]}
                maxLength={100_000}
                onChange={(event) => setValues({ ...values, [key]: event.target.value })}
              />
            ) : (
              <input
                type="text"
                value={values[key]}
                maxLength={2_000}
                autoFocus={index === 0}
                onChange={(event) => setValues({ ...values, [key]: event.target.value })}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') save()
                }}
              />
            )}
          </label>
        ))}
      </fieldset>

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

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={close}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onClick={save}>
          {t('document.common.apply')}
        </button>
      </div>
    </div>
  )
}
