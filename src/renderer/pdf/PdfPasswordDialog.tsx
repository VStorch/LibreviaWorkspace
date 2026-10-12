import { useState } from 'react'
import { DialogActions } from '../components/DialogActions.js'
import { useT } from '../i18n.js'

/** The password goes from here straight to the viewer frame; nothing else sees or keeps it. */
export function PdfPasswordDialog({
  retry,
  onAnswer,
  onCancel,
}: {
  readonly retry: boolean
  readonly onAnswer: (password: string) => void
  readonly onCancel: () => void
}): React.JSX.Element {
  const t = useT()
  const [password, setPassword] = useState('')
  const answer = (): void => {
    if (password !== '') onAnswer(password)
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('shell.pdf.passwordTitle')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onCancel()
      }}
    >
      <label className="popover__field">
        <span>{t('shell.pdf.passwordLabel')}</span>
        <input
          type="password"
          value={password}
          autoFocus
          autoComplete="off"
          onChange={(event) => setPassword(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') answer()
          }}
        />
      </label>
      {retry && (
        <p className="popover__error" role="alert">
          {t('shell.pdf.passwordWrong')}
        </p>
      )}
      <p className="popover__hint">{t('shell.pdf.passwordHint')}</p>
      <DialogActions
        confirmLabel={t('shell.pdf.passwordOpen')}
        onConfirm={answer}
        onCancel={onCancel}
        disabled={password === ''}
      />
    </div>
  )
}
