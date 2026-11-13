import { useState } from 'react'
import { LANGUAGES, LANGUAGE_NAMES, type Language, type MessageKey } from '@shared/i18n/index.js'
import { MAX_NAME_LENGTH } from '@shared/limits.js'
import { SAVE_FORMATS, SaveFormat, Theme, type EditorPreferences } from '@shared/types.js'
import { useT } from '../i18n.js'
import { setPreference, usePreferences } from '../state/preferences.js'
import { DialogActions } from './DialogActions.js'
import { ShortcutsDialog } from './ShortcutsDialog.js'

type OptionValues = Pick<
  EditorPreferences,
  'saveFormat' | 'authorName' | 'language' | 'theme' | 'spellcheck' | 'typography'
>

const SAVE_FORMAT_LABELS: Record<SaveFormat, MessageKey> = {
  [SaveFormat.Ask]: 'shell.options.saveFormat.ask',
  [SaveFormat.Office]: 'shell.options.saveFormat.office',
  [SaveFormat.LibreOffice]: 'shell.options.saveFormat.libreOffice',
  [SaveFormat.Librevia]: 'shell.options.saveFormat.librevia',
}

const THEME_LABELS: Record<Theme, MessageKey> = {
  [Theme.System]: 'view.theme.system',
  [Theme.Light]: 'view.theme.light',
  [Theme.Dark]: 'view.theme.dark',
}

/** Tools → Options: the preferences the menu has no room for, next to the ones it already toggles. */
export function OptionsDialog({ onClose }: { readonly onClose: () => void }): React.JSX.Element {
  const t = useT()
  const stored = usePreferences((state) => state.preferences)
  const [values, setValues] = useState<OptionValues>(() => ({
    saveFormat: stored.saveFormat,
    authorName: stored.authorName,
    language: stored.language,
    theme: stored.theme,
    spellcheck: stored.spellcheck,
    typography: stored.typography,
  }))
  const [shortcuts, setShortcuts] = useState(false)
  const change = (patch: Partial<OptionValues>): void => setValues((current) => ({ ...current, ...patch }))

  function apply(): void {
    void setPreference({ ...values, authorName: values.authorName.trim() })
    onClose()
  }

  if (shortcuts) return <ShortcutsDialog onClose={() => setShortcuts(false)} />
  return (
    <div
      className="popover popover--wide options"
      role="dialog"
      aria-label={t('shell.options.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <FileOptions values={values} onChange={change} />
      <AppearanceOptions values={values} onChange={change} />
      <EditingOptions values={values} onChange={change} />
      <fieldset className="options__group">
        <legend>{t('shell.options.keyboard')}</legend>
        <button type="button" className="btn" onClick={() => setShortcuts(true)}>
          {t('shell.options.customizeShortcuts')}
        </button>
      </fieldset>
      <DialogActions confirmLabel={t('document.common.apply')} onConfirm={apply} onCancel={onClose} />
    </div>
  )
}

interface SectionProps {
  readonly values: OptionValues
  readonly onChange: (patch: Partial<OptionValues>) => void
}

function FileOptions({ values, onChange }: SectionProps): React.JSX.Element {
  const t = useT()
  return (
    <fieldset className="options__group">
      <legend>{t('shell.options.files')}</legend>
      <label className="popover__field">
        <span>{t('shell.options.saveFormat')}</span>
        <select
          autoFocus
          value={values.saveFormat}
          onChange={(event) => onChange({ saveFormat: event.target.value as SaveFormat })}
        >
          {SAVE_FORMATS.map((format) => (
            <option key={format} value={format}>
              {t(SAVE_FORMAT_LABELS[format])}
            </option>
          ))}
        </select>
      </label>
      <p className="popover__hint">{t('shell.options.saveFormatHint')}</p>
    </fieldset>
  )
}

function AppearanceOptions({ values, onChange }: SectionProps): React.JSX.Element {
  const t = useT()
  return (
    <fieldset className="options__group">
      <legend>{t('shell.options.appearance')}</legend>
      <div className="popover__row">
        <label className="popover__field">
          <span>{t('view.language')}</span>
          <select
            value={values.language}
            onChange={(event) => onChange({ language: event.target.value as Language })}
          >
            {LANGUAGES.map((language) => (
              <option key={language} value={language}>
                {LANGUAGE_NAMES[language]}
              </option>
            ))}
          </select>
        </label>
        <label className="popover__field">
          <span>{t('view.theme')}</span>
          <select value={values.theme} onChange={(event) => onChange({ theme: event.target.value as Theme })}>
            {Object.values(Theme).map((theme) => (
              <option key={theme} value={theme}>
                {t(THEME_LABELS[theme])}
              </option>
            ))}
          </select>
        </label>
      </div>
    </fieldset>
  )
}

function EditingOptions({ values, onChange }: SectionProps): React.JSX.Element {
  const t = useT()
  return (
    <fieldset className="options__group">
      <legend>{t('shell.options.editing')}</legend>
      <label className="popover__field">
        <span>{t('comments.author.title')}</span>
        <input
          type="text"
          value={values.authorName}
          maxLength={MAX_NAME_LENGTH}
          onChange={(event) => onChange({ authorName: event.target.value })}
        />
      </label>
      <label className="popover__check">
        <input
          type="checkbox"
          checked={values.spellcheck}
          onChange={(event) => onChange({ spellcheck: event.target.checked })}
        />
        {t('menu.tools.spellcheck')}
      </label>
      <label className="popover__check">
        <input
          type="checkbox"
          checked={values.typography}
          onChange={(event) => onChange({ typography: event.target.checked })}
        />
        {t('menu.tools.typography')}
      </label>
    </fieldset>
  )
}
