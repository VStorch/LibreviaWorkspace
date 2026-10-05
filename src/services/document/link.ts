import { ALLOWED_EXTERNAL_PROTOCOLS } from '@shared/constants.js'

/**
 * `javascript:` in a link is disguised code. Main applies the same scheme list before opening; this
 * only avoids saving an invalid address.
 */
export function normalizeLinkUrl(input: string): string | null {
  const trimmed = input.trim()
  if (trimmed.length === 0) return null

  // An address typed without a scheme is the common case: "empresa.com.br".
  const candidate = /^[a-z][a-z0-9+.-]*:/i.test(trimmed) ? trimmed : `https://${trimmed}`

  let parsed: URL
  try {
    parsed = new URL(candidate)
  } catch {
    return null
  }

  if (!(ALLOWED_EXTERNAL_PROTOCOLS as readonly string[]).includes(parsed.protocol)) return null
  if (parsed.protocol !== 'mailto:' && parsed.hostname === '') return null

  return parsed.toString()
}
