import { z } from 'zod'
import { isMessageKey } from '@shared/i18n/index.js'
import { t } from '../i18n.js'

const SIDECAR_KEY_PREFIX = 'sidecar.'

/** The cut of `ipc.ts`: without it, a pathological document would fail to open because of the warning. */
const MAX_NOTICES = 50
const MAX_NOTICE_LENGTH = 300

function knownKey(key: string) {
  return key.startsWith(SIDECAR_KEY_PREFIX) && isMessageKey(key) ? key : null
}

/** A key this build does not have comes from an older or newer sidecar: the generic sentence. */
export function sidecarMessage(key: string): string {
  return t(knownKey(key) ?? 'sidecar.error.unexpected')
}

export interface SidecarNotice {
  readonly key: string
  readonly value?: string | undefined
}

export function sidecarNotice({ key, value }: SidecarNotice): string {
  const known = knownKey(key)
  if (known === null) return t('errors.sidecar.unknownNotice')
  return t(known, value === undefined ? undefined : { value })
}

const notices = z
  .array(z.object({ key: z.string(), value: z.string().optional() }))
  .default([])
  .transform((list) => list.slice(0, MAX_NOTICES).map((notice) => sidecarNotice(notice).slice(0, MAX_NOTICE_LENGTH)))

/** The sidecar names each item by key; main hands the screen the sentences. */
export const sidecarInventorySchema = z.object({
  invisible: notices,
  lost: notices,
  structural: notices,
})
