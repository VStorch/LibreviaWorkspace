import { DocumentKind } from '@shared/types.js'
import { useWorkspace } from '../state/workspace.js'

/**
 * Havia trabalho não salvo quando a sessão anterior terminou.
 *
 * Faixa com duas ações, e não modal, pela mesma razão dos outros avisos — mas
 * aqui há uma segunda razão, mais forte: enquanto esta faixa estiver na tela o
 * autosave **não escreve**, então ignorá-la não custa nada ao usuário. O
 * rascunho só desaparece quando ele mesmo diz para descartar.
 *
 * O texto diz de quando é o rascunho porque essa é a informação que decide: um
 * rascunho de dois minutos atrás quase sempre vale mais que o arquivo em disco,
 * e um de três semanas quase nunca.
 */
export function RecoveryBanner(): React.JSX.Element | null {
  const draft = useWorkspace((state) => state.pendingDraft)
  const recover = useWorkspace((state) => state.recoverDraft)
  const dismiss = useWorkspace((state) => state.dismissDraft)

  if (draft === null) return null

  const what = draft.kind === DocumentKind.Spreadsheet ? 'uma planilha' : 'um documento'

  return (
    <div className="banner banner--recovery" role="alert">
      <div className="banner__text">
        <strong>O aplicativo fechou com trabalho não salvo.</strong>
        <span className="banner__detail">
          Há {what} de “{draft.name}”, guardado {when(draft.savedAt)}. Recuperar traz o conteúdo de volta para
          a tela sem gravar em nada — você decide onde salvar.
        </span>
      </div>
      <div className="banner__actions">
        <button type="button" className="banner__action" onClick={() => void recover()}>
          Recuperar
        </button>
        <button type="button" className="banner__action banner__action--quiet" onClick={() => void dismiss()}>
          Descartar
        </button>
      </div>
    </div>
  )
}

/**
 * "há 3 minutos", "ontem às 17:42".
 *
 * Tempo relativo perto e data absoluta longe: "há 26 dias" não ajuda ninguém a
 * decidir, e "às 17:42" de hoje de manhã tampouco.
 */
function when(savedAt: number): string {
  const minutes = Math.round((Date.now() - savedAt) / 60_000)
  if (minutes < 1) return 'agora há pouco'
  if (minutes < 60) return `há ${minutes} ${minutes === 1 ? 'minuto' : 'minutos'}`

  const hours = Math.round(minutes / 60)
  if (hours < 24) return `há ${hours} ${hours === 1 ? 'hora' : 'horas'}`

  return `em ${new Date(savedAt).toLocaleString('pt-BR', {
    day: '2-digit',
    month: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  })}`
}
