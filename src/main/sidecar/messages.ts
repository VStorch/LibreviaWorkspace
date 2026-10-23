import { isMessageKey } from '@shared/i18n/index.js'
import { t } from '../i18n.js'

const SIDECAR_KEY_PREFIX = 'sidecar.'

/** A key this build does not have comes from an older or newer sidecar: the generic sentence. */
export function sidecarMessage(key: string): string {
  return key.startsWith(SIDECAR_KEY_PREFIX) && isMessageKey(key) ? t(key) : t('sidecar.error.unexpected')
}
