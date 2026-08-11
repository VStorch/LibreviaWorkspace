import { useWorkspace } from '../state/workspace.js'

/**
 * Erros são mostrados como faixa, não como diálogo modal.
 *
 * Falha de leitura ou gravação quase sempre é recuperável (rede fora do ar,
 * arquivo movido). Interromper com um modal força uma resposta que o usuário
 * ainda não tem; a faixa informa e deixa continuar trabalhando.
 */
export function ErrorBanner(): React.JSX.Element | null {
  const error = useWorkspace((state) => state.error)
  const dismiss = useWorkspace((state) => state.dismissError)

  if (error === null) return null

  return (
    <div className="banner" role="alert">
      <div className="banner__text">
        <strong>{error.message}</strong>
        {error.detail !== undefined && <span className="banner__detail">{error.detail}</span>}
      </div>
      <button type="button" className="banner__close" onClick={dismiss} aria-label="Dispensar aviso">
        ✕
      </button>
    </div>
  )
}
