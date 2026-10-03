import { DocumentKind } from '@shared/types.js'
import { Language, type MessageKey, type Vars } from '@shared/i18n/index.js'
import { useLanguage, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'

/**
 * Enquanto a faixa estiver na tela o autosave **não escreve**: ignorá-la não
 * custa nada. A data é o que decide entre o rascunho e o arquivo em disco.
 */
export function RecoveryBanner(): React.JSX.Element | null {
  const t = useT()
  const language = useLanguage()
  const draft = useWorkspace((state) => state.pendingDraft)
  const recover = useWorkspace((state) => state.recoverDraft)
  const dismiss = useWorkspace((state) => state.dismissDraft)

  if (draft === null) return null

  const what =
    draft.kind === DocumentKind.Spreadsheet ? t('shell.recovery.spreadsheet') : t('shell.recovery.document')

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

/** Tempo relativo perto e data absoluta longe: "há 26 dias" não ajuda a decidir. */
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
