import { ALLOWED_EXTERNAL_PROTOCOLS } from '@shared/constants.js'
import { safeHref, type Mark } from './export-common.js'

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

/**
 * An imported document's links, without those `safeHref` refuses: the text stays, unlinked. The
 * document comes from a file, so its shape is not trusted.
 */
export function withoutUnsafeLinks(node: unknown): unknown {
  if (Array.isArray(node)) return node.map(withoutUnsafeLinks)
  if (typeof node !== 'object' || node === null) return node
  const { marks, content, ...rest } = node as Record<string, unknown>
  const kept = Array.isArray(marks) ? marks.filter((mark) => !isUnsafeLink(mark)) : marks
  return {
    ...rest,
    ...(kept === undefined || (Array.isArray(kept) && kept.length === 0) ? {} : { marks: kept }),
    ...(content === undefined ? {} : { content: withoutUnsafeLinks(content) }),
  }
}

function isUnsafeLink(mark: unknown): boolean {
  if (typeof mark !== 'object' || mark === null) return false
  const candidate = mark as Mark
  return candidate.type === 'link' && safeHref(candidate) === null
}
