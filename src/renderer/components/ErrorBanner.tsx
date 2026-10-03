import { useWorkspace } from '../state/workspace.js'
import { useT } from '../i18n.js'

/** Faixa, e não modal: a falha quase sempre é recuperável, e o modal forçaria uma resposta que a pessoa ainda não tem. */
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
