import { useWorkspace } from '../state/workspace.js'

/**
 * O arquivo abriu travado, e a faixa diz por quê e como destravar.
 *
 * A proteção é contra o caso concreto: a gravação é cirúrgica, então comentário,
 * revisão e nota voltam intactos **desde que você não edite o bloco que os
 * ancora**. Quem só precisa ler não corre risco nenhum; quem precisa editar
 * clica uma vez e assume o risco sabendo qual é.
 *
 * Por isso é padrão e não cadeado, e por isso a lista do que está em jogo
 * aparece junto do botão: um aviso que não diz o que se perde é um aviso que se
 * fecha sem ler.
 */
export function ReadOnlyBanner(): React.JSX.Element | null {
  const readOnly = useWorkspace((state) => state.readOnly)
  const notice = useWorkspace((state) => state.notice)
  const allowEditing = useWorkspace((state) => state.allowEditing)

  if (!readOnly) return null

  const reasons = notice?.structural ?? []

  return (
    <div className="banner banner--readonly" role="status">
      <div className="banner__text">
        <strong>Aberto somente para leitura.</strong>
        <span className="banner__detail">
          {reasons.length > 0
            ? `Este arquivo tem ${reasons.join(', ')} — que o editor não mostra. Tudo isso volta intacto ao salvar, menos o que estiver no trecho que você editar.`
            : 'Este arquivo tem recursos que o editor não mostra e que podem se perder ao editar.'}
        </span>
      </div>
      <div className="banner__actions">
        <button type="button" className="banner__action" onClick={allowEditing}>
          Editar mesmo assim
        </button>
      </div>
    </div>
  )
}
