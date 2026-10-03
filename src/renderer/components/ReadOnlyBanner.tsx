import { useWorkspace } from '../state/workspace.js'
import { useT } from '../i18n.js'

/**
 * Comentário, revisão e nota voltam intactos desde que não se edite o bloco que
 * os ancora: é padrão, e não cadeado, e a lista do que está em jogo fica junto do
 * botão. "Não reproduz por inteiro", e não "não mostra": o texto das caixas
 * aparece, a moldura e a posição não.
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
