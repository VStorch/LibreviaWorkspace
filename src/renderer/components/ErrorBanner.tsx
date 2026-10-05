import { useWorkspace } from '../state/workspace.js'
import { useT } from '../i18n.js'

/**
 * A banner, not a modal: the failure is almost always recoverable, and a modal would force an
 * answer the user does not have yet.
 */
export function ErrorBanner(): React.JSX.Element | null {
  const t = useT()
  const error = useWorkspace((state) => state.error)
  const dismiss = useWorkspace((state) => state.dismissError)

  if (error === null) return null

  return (
    <div className="banner" role="alert">
      <div className="banner__text">
        <strong>{error.message}</strong>
        {error.detail !== undefined && <span className="banner__detail">{error.detail}</span>}
      </div>
      <button
        type="button"
        className="banner__close"
        onClick={dismiss}
        aria-label={t('shell.banner.dismiss')}
      >
        ✕
      </button>
    </div>
  )
}
