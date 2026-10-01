import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { useT } from '../i18n.js'
import { currentPreferences, setPreference } from '../state/preferences.js'

/**
 * Ferramentas → Nome do autor: quem assina os comentários novos, e as iniciais
 * que saem dele. Os que já existem guardam o autor com que foram escritos, como
 * no Word.
 */
export function AuthorNameDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const [name, setName] = useState(() => currentPreferences().authorName)

  function save(): void {
    void setPreference({ authorName: name.trim() })
    onClose()
    requestAnimationFrame(() => editor.view.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('comments.author.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <label className="popover__field">
        <span>{t('comments.author.label')}</span>
        <input
          type="text"
          value={name}
          maxLength={200}
          autoFocus
          onChange={(event) => setName(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') save()
          }}
        />
      </label>
      <p className="popover__hint">{t('comments.author.hint')}</p>
      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onClick={save}>
          {t('comments.action.save')}
        </button>
      </div>
    </div>
  )
}
