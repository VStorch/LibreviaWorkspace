import { t } from '../i18n.js'
import type { SetWorkspace } from './context.js'

/** Saying nothing would make "Save" or "Print" look broken. */
export function refuseViewOnly(set: SetWorkspace): false {
  set({ error: { code: 'INTERNAL', message: t('shell.pdf.viewOnly') } })
  return false
}
