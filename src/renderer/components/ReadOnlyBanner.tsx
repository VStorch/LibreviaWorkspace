import { useWorkspace } from '../state/workspace.js'
import { useT } from '../i18n.js'

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
 *
 * "Não reproduz por inteiro", e não "não mostra": desde que o leitor passou a
 * trazer o texto de dentro das caixas, dizer que elas não aparecem seria falso
 * na metade dos casos — o texto aparece, a moldura e a posição não. Um aviso
 * que o usuário consegue desmentir olhando a tela é pior que nenhum.
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
