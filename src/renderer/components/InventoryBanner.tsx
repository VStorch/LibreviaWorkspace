import { useWorkspace } from '../state/workspace.js'
import { useT } from '../i18n.js'

/**
 * The two lists stay **apart**: one says "it exists and you cannot see it", the other "it will
 * disappear".
 */
export function InventoryBanner(): React.JSX.Element | null {
  const t = useT()
  const notice = useWorkspace((state) => state.notice)
  const savedLoss = useWorkspace((state) => state.savedLoss)
  const readOnly = useWorkspace((state) => state.readOnly)
  const dismiss = useWorkspace((state) => state.dismissNotice)

  // What the save just lost comes first, and alone.
  if (savedLoss !== null && savedLoss.length > 0) {
    return (
      <div className="banner banner--notice" role="status">
        <div className="banner__text">
          <strong>{t('shell.banner.savedLoss')}</strong>
          <span className="banner__detail">{savedLoss.join('; ')}</span>
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

  if (notice === null) return null

  // With the read-only banner on screen, the structural items are already named; once unlocked, the
  // full list comes back.
  const invisible = readOnly
    ? notice.invisible.filter((item) => !notice.structural.includes(item))
    : notice.invisible

  if (invisible.length === 0 && notice.lost.length === 0) return null

  return (
    <div className="banner banner--notice" role="status">
      <div className="banner__text">
        {notice.lost.length > 0 && (
          <>
            <strong>{t('shell.banner.willBeLost')}</strong>
            <span className="banner__detail">{notice.lost.join('; ')}</span>
          </>
        )}
        {invisible.length > 0 && (
          <>
            <strong>
              {notice.lost.length > 0 ? t('shell.banner.stillInFileAnd') : t('shell.banner.stillInFileDoc')}
            </strong>
            <span className="banner__detail">{invisible.join('; ')}</span>
          </>
        )}
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
