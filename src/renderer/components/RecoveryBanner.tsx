import { DocumentKind } from '@shared/types.js'
import { Language, type MessageKey, type Vars } from '@shared/i18n/index.js'
import { useLanguage, useT } from '../i18n.js'
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
  const t = useT()
  const language = useLanguage()
  const draft = useWorkspace((state) => state.pendingDraft)
  const recover = useWorkspace((state) => state.recoverDraft)
  const dismiss = useWorkspace((state) => state.dismissDraft)

  if (draft === null) return null

  const what = draft.kind === DocumentKind.Spreadsheet
    ? t('shell.recovery.spreadsheet')
    : t('shell.recovery.document')

  return (
    <div className="banner banner--recovery" role="alert">
      <div className="banner__text">
        <strong>{t('shell.recovery.title')}</strong>
        <span className="banner__detail">
          {t('shell.recovery.detail', {
            what,
            name: draft.name,
            time: when(draft.savedAt, t, language),
          })}
        </span>
      </div>
      <div className="banner__actions">
        <button type="button" className="banner__action" onClick={() => void recover()}>
          {t('shell.recovery.recover')}
        </button>
        <button type="button" className="banner__action banner__action--quiet" onClick={() => void dismiss()}>
          {t('shell.recovery.discard')}
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
function when(savedAt: number, t: (key: MessageKey, vars?: Vars) => string, language: Language): string {
  const minutes = Math.round((Date.now() - savedAt) / 60_000)
  if (minutes < 1) return t('shell.recovery.justNow')
  if (minutes < 60) return t('shell.recovery.minutesAgo', { count: minutes })

  const hours = Math.round(minutes / 60)
  if (hours < 24) return t('shell.recovery.hoursAgo', { count: hours })

  const locale = language === Language.English ? 'en-US' : 'pt-BR'
  const dateStr = new Date(savedAt).toLocaleString(locale, {
    day: '2-digit',
    month: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  })
  return t('shell.recovery.onDate', { date: dateStr })
}
