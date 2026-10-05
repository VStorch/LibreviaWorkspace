import { useWorkspace } from '../state/workspace.js'
import { useT } from '../i18n.js'

/**
 * Comments, revisions and notes come back intact as long as the block anchoring them is not edited:
 * a default, not a lock, and the list of what is at stake sits next to the button. "Does not fully
 * reproduce", not "does not show": text box contents appear, their frame and position do not.
 */
export function ReadOnlyBanner(): React.JSX.Element | null {
  const t = useT()
  const readOnly = useWorkspace((state) => state.readOnly)
  const notice = useWorkspace((state) => state.notice)
  const allowEditing = useWorkspace((state) => state.allowEditing)

  if (!readOnly) return null

  const reasons = notice?.structural ?? []

  return (
    <div className="banner banner--readonly" role="status">
      <div className="banner__text">
        <strong>{t('shell.banner.readOnly')}</strong>
        <span className="banner__detail">
          {reasons.length > 0
            ? t('shell.banner.readOnlyReasons', { reasons: reasons.join(', ') })
            : t('shell.banner.readOnlyGeneric')}
        </span>
      </div>
      <div className="banner__actions">
        <button type="button" className="banner__action" onClick={allowEditing}>
          {t('shell.banner.editAnyway')}
        </button>
      </div>
    </div>
  )
}
