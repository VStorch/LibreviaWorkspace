import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { normalizeLinkUrl } from '@services/document/link.js'
import { useT } from '../../i18n.js'

interface LinkDialogProps {
  readonly editor: Editor
  readonly onClose: () => void
}

export function LinkDialog({ editor, onClose }: LinkDialogProps): React.JSX.Element {
  const t = useT()
  const existing = String(editor.getAttributes('link')['href'] ?? '')
  const [value, setValue] = useState(existing)
  const [rejected, setRejected] = useState(false)

  function apply(): void {
    const normalized = normalizeLinkUrl(value)
    if (normalized === null) {
      setRejected(true)
      return
    }

    editor.chain().focus().extendMarkRange('link').setLink({ href: normalized }).run()
    onClose()
  }

  function remove(): void {
    editor.chain().focus().extendMarkRange('link').unsetLink().run()
    onClose()
  }

  return (
    <div className="popover" role="dialog" aria-label={t('document.linkDialog.title')}>
      <label className="popover__field">
        <span>{t('document.linkDialog.address')}</span>
        <input
          type="text"
          value={value}
          autoFocus
          placeholder={t('document.linkDialog.placeholder')}
          onChange={(event) => {
            setValue(event.target.value)
            setRejected(false)
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter') apply()
            if (event.key === 'Escape') onClose()
          }}
        />
      </label>

      {rejected && <p className="popover__error">{t('document.linkDialog.invalidAddress')}</p>}

      <div className="popover__actions">
        {existing !== '' && (
          <button type="button" className="btn" onClick={remove}>
            {t('document.linkDialog.remove')}
          </button>
        )}
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onClick={apply}>
          {t('document.common.apply')}
        </button>
      </div>
    </div>
  )
}
