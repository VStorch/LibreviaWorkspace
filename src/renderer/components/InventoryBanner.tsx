import { useWorkspace } from '../state/workspace.js'

/**
 * O que o documento tem e o editor não mostra — ou não vai preservar.
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
  const notice = useWorkspace((state) => state.notice)
  const dismiss = useWorkspace((state) => state.dismissNotice)

  if (notice === null) return null

  return (
    <div className="banner banner--notice" role="status">
      <div className="banner__text">
        {notice.lost.length > 0 && (
          <>
            <strong>Ao salvar, isto será perdido:</strong>
            <span className="banner__detail">{notice.lost.join('; ')}</span>
          </>
        )}
        {notice.invisible.length > 0 && (
          <>
            <strong>
              {notice.lost.length > 0
                ? 'E isto continua no arquivo, mas não aparece aqui:'
                : 'Este documento tem recursos que continuam no arquivo, mas não aparecem aqui:'}
            </strong>
            <span className="banner__detail">{notice.invisible.join('; ')}</span>
          </>
        )}
      </div>
      <button type="button" className="banner__close" onClick={dismiss} aria-label="Dispensar aviso">
        ✕
      </button>
    </div>
  )
}
