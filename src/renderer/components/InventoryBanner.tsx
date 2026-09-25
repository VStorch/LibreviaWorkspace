import { useWorkspace } from '../state/workspace.js'
import { useT } from '../i18n.js'

/**
 * O que o documento tem e o editor não reproduz por inteiro — ou não vai
 * preservar.
 *
 * Faixa, e não modal, pela mesma razão do `ErrorBanner`: o arquivo abriu e o
 * usuário pode trabalhar. Um modal aqui treinaria a mão a clicar "OK" sem ler,
 * e aí o dia em que o aviso for grave ele também passa direto.
 *
 * As duas listas ficam **separadas na tela**, e não numa só, porque são
 * problemas diferentes: uma diz "existe e você não vê", a outra diz "vai
 * sumir". Ver docs/02-docx-cirurgico.md.
 */
export function InventoryBanner(): React.JSX.Element | null {
  const t = useT()
  const notice = useWorkspace((state) => state.notice)
  const savedLoss = useWorkspace((state) => state.savedLoss)
  const readOnly = useWorkspace((state) => state.readOnly)
  const dismiss = useWorkspace((state) => state.dismissNotice)

  // O que a gravação acabou de perder vem primeiro, e sozinho: é o aviso de
  // agora, e o da abertura já foi lido quando o arquivo abriu.
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

  // Enquanto a faixa de somente leitura está na tela, ela já nomeia o que é
  // estrutural. Repetir aqui empilharia dois avisos dizendo a mesma coisa, e
  // dois avisos iguais valem menos que um. Ao liberar a edição a lista volta
  // inteira — que é justamente quando ela passa a importar.
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
