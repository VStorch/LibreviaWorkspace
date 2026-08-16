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
  const readOnly = useWorkspace((state) => state.readOnly)
  const dismiss = useWorkspace((state) => state.dismissNotice)

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
            <strong>Ao salvar, isto será perdido:</strong>
            <span className="banner__detail">{notice.lost.join('; ')}</span>
          </>
        )}
        {invisible.length > 0 && (
          <>
            <strong>
              {notice.lost.length > 0
                ? 'E isto continua no arquivo, mas não aparece aqui:'
                : 'Este documento tem recursos que continuam no arquivo, mas não aparecem aqui:'}
            </strong>
            <span className="banner__detail">{invisible.join('; ')}</span>
          </>
        )}
      </div>
      <button type="button" className="banner__close" onClick={dismiss} aria-label="Dispensar aviso">
        ✕
      </button>
    </div>
  )
}
